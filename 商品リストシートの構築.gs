/*******************************************************
 * 商品リストシート構築 V1
 *
 * 目的：
 * - 一度でも入庫・受け入れされた商品だけを表示
 * - VaMASTERの商品情報を商品リストへ展開
 * - SKUから064単位で在庫をサイズ別集計
 * - 全在庫と倉庫在庫だけマトリックス表示
 * - 在庫あり商品を上、在庫なし商品を下に並べる
 *
 * シート名：
 * - VaMASTER
 * - SKU
 * - 商品リスト
 *******************************************************/

const PRODUCT_LIST_CONFIG = {
  SOURCE_VAMASTER: "VaMASTER",
  SOURCE_SKU: "SKU",
  TARGET_SHEET: "商品リスト",

  HEADER_ROW: 6,
  DATA_START_ROW: 7,

  // 商品リストの固定列
  TARGET: {
    TOTAL_COLS: 51, // AY列まで

    // AI:AO 全在庫
    ALL_STOCK_START_COL: 35, // AI
    ALL_STOCK_SUM_COL: 41,   // AO

    // AP:AV 倉庫在庫
    WH_STOCK_START_COL: 42,  // AP
    WH_STOCK_SUM_COL: 48,    // AV

    // AW:AY 管理列
    SORT_COL: 49,            // AW 4107_並び順
    DISPLAY_TYPE_COL: 50,    // AX 4106_表示区分
    SALES_STATUS_COL: 51     // AY 4040_販売ステータス
  },

  SIZE_ORDER: ["XS", "S", "M", "L", "XL", "F"],

  // 商品リスト側で手入力・編集される可能性がある項目ID
  // 再構築しても、064が一致すれば復元します。
  MANUAL_KEEP_IDS: [
    "15",    // 卸価格(¥) ※式や手入力を残したい場合
    "17",    // 商品名(TT)
    "1000",  // TikTok価格
    "1001",  // 移動販売価格
    "1002",  // レンタル価格
    "21",    // 備考
    "20",    // 商品説明
    "23",    // Category
    "24",    // Collection
    "4021",  // キャッチコピー
    "4022",  // 推しポイント
    "4040"   // 販売ステータス
  ],

  // copyToで維持したい項目ID
  // URLチップ・スマートチップ・リッチ表示がある可能性があるもの
  COPY_TO_IDS: ["03", "20", "21", "22"]
};


/*******************************************************
 * メイン関数
 *******************************************************/

function generateProductListSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const vaSheet = ss.getSheetByName(PRODUCT_LIST_CONFIG.SOURCE_VAMASTER);
  const skuSheet = ss.getSheetByName(PRODUCT_LIST_CONFIG.SOURCE_SKU);
  const targetSheet = ss.getSheetByName(PRODUCT_LIST_CONFIG.TARGET_SHEET);

  if (!vaSheet || !skuSheet || !targetSheet) {
    Browser.msgBox("VaMASTER / SKU / 商品リスト のいずれかのシートが見つかりません。");
    return;
  }

  const confirm = Browser.msgBox(
    "商品リストシート構築",
    "商品リストシートを再構築します。\n既存の販売価格・説明文・ステータスなどは064コードをキーに復元します。\n\n実行しますか？",
    Browser.Buttons.YES_NO
  );

  if (confirm !== "yes") return;

  const vaColMap = getProductListColMap_(vaSheet);
  const skuColMap = getProductListColMap_(skuSheet);
  const tgtColMap = getProductListColMap_(targetSheet);

  if (!vaColMap["064"]) {
    Browser.msgBox("VaMASTERに 064_ 商品コード列が見つかりません。");
    return;
  }

  if (!skuColMap["064"]) {
    Browser.msgBox("SKUシートに 064_ 商品コード列が見つかりません。");
    return;
  }

  if (!skuColMap["09"]) {
    Browser.msgBox("SKUシートに 09_ サイズ列が見つかりません。");
    return;
  }

  if (!tgtColMap["064"]) {
    Browser.msgBox("商品リストシートに 064_ 商品コード列が見つかりません。");
    return;
  }

  // ==========================================
  // Step 1: 商品リストの既存入力バックアップ
  // ==========================================

  const backupMap = new Map();
  const targetLastRow = targetSheet.getLastRow();

  if (targetLastRow >= PRODUCT_LIST_CONFIG.DATA_START_ROW) {
    const existingValues = targetSheet.getRange(
      PRODUCT_LIST_CONFIG.DATA_START_ROW,
      1,
      targetLastRow - PRODUCT_LIST_CONFIG.DATA_START_ROW + 1,
      PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS
    ).getValues();

    existingValues.forEach(row => {
      const key064 = String(row[tgtColMap["064"] - 1] || "").trim();
      if (!key064) return;

      const saved = {};

      PRODUCT_LIST_CONFIG.MANUAL_KEEP_IDS.forEach(id => {
        if (!tgtColMap[id]) return;

        const col = tgtColMap[id];
        saved[col] = row[col - 1];
      });

      // 固定列側の販売ステータスも一応保存
      saved[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL] =
        row[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL - 1];

      backupMap.set(key064, saved);
    });
  }

  // ==========================================
  // Step 2: SKUから表示対象と在庫集計を作る
  // ==========================================

  const skuLastRow = skuSheet.getLastRow();
  const skuLastCol = skuSheet.getLastColumn();

  if (skuLastRow < PRODUCT_LIST_CONFIG.DATA_START_ROW) {
    Browser.msgBox("SKUシートにデータがありません。");
    return;
  }

  const skuValues = skuSheet.getRange(
    PRODUCT_LIST_CONFIG.DATA_START_ROW,
    1,
    skuLastRow - PRODUCT_LIST_CONFIG.DATA_START_ROW + 1,
    skuLastCol
  ).getValues();

  // 表示対象判定に使う累計入庫系ID
  const intakeIds = ["50", "54", "58", "71", "75"];

  // 現在庫ID
  const allStockId = "2001";
  const warehouseStockId = "52";

  const handled064Set = new Set();

  const stockMap = new Map();
  // stockMap[064] = {
  //   all: {XS:0,S:0,...},
  //   wh:  {XS:0,S:0,...},
  //   allTotal: 0,
  //   whTotal: 0
  // }

  skuValues.forEach(row => {
    const key064 = String(row[skuColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const size = normalizeProductListSize_(row[skuColMap["09"] - 1]);
    if (!PRODUCT_LIST_CONFIG.SIZE_ORDER.includes(size)) return;

    const intakeTotal = intakeIds.reduce((sum, id) => {
      if (!skuColMap[id]) return sum;
      return sum + toProductListNumber_(row[skuColMap[id] - 1]);
    }, 0);

    if (intakeTotal > 0) {
      handled064Set.add(key064);
    }

    if (!stockMap.has(key064)) {
      stockMap.set(key064, {
        all: createProductListSizeObject_(),
        wh: createProductListSizeObject_(),
        allTotal: 0,
        whTotal: 0
      });
    }

    const info = stockMap.get(key064);

    let allQty = 0;
    if (skuColMap[allStockId]) {
      allQty = toProductListNumber_(row[skuColMap[allStockId] - 1]);
    } else {
      // 2001がない場合の保険
      ["52", "56", "60", "73", "76"].forEach(id => {
        if (skuColMap[id]) {
          allQty += toProductListNumber_(row[skuColMap[id] - 1]);
        }
      });
    }

    const whQty = skuColMap[warehouseStockId]
      ? toProductListNumber_(row[skuColMap[warehouseStockId] - 1])
      : 0;

    info.all[size] += allQty;
    info.wh[size] += whQty;
    info.allTotal += allQty;
    info.whTotal += whQty;
  });

  if (handled064Set.size === 0) {
    Browser.msgBox("一度でも入庫・受け入れされた商品が見つかりませんでした。");
    return;
  }

  // ==========================================
  // Step 3: VaMASTERを読む
  // ==========================================

  const vaLastRow = vaSheet.getLastRow();
  const vaLastCol = vaSheet.getLastColumn();

  if (vaLastRow < PRODUCT_LIST_CONFIG.DATA_START_ROW) {
    Browser.msgBox("VaMASTERにデータがありません。");
    return;
  }

  const vaValues = vaSheet.getRange(
    PRODUCT_LIST_CONFIG.DATA_START_ROW,
    1,
    vaLastRow - PRODUCT_LIST_CONFIG.DATA_START_ROW + 1,
    vaLastCol
  ).getValues();

  // VaMASTERの064検索Map
  const vaBy064 = new Map();

  vaValues.forEach((row, idx) => {
    const key064 = String(row[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    vaBy064.set(key064, {
      row,
      rowNumber: PRODUCT_LIST_CONFIG.DATA_START_ROW + idx
    });
  });

  // ==========================================
  // Step 4: 商品リストの初期化
  // ==========================================

  if (targetLastRow >= PRODUCT_LIST_CONFIG.DATA_START_ROW) {
    targetSheet.getRange(
      PRODUCT_LIST_CONFIG.DATA_START_ROW,
      1,
      targetLastRow - PRODUCT_LIST_CONFIG.DATA_START_ROW + 1,
      PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS
    )
      .clearContent()
      .setBackground(null);
  }

  // ==========================================
  // Step 5: 出力データ作成
  // ==========================================

  const outputRows = [];
  const outputBgs = [];
  const copyTasks = [];

  vaValues.forEach((vaRow, vaIndex) => {
    const key064 = String(vaRow[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    // 一度も入庫・受け入れされていない商品は出さない
    if (!handled064Set.has(key064)) return;

    const rowData = new Array(PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS).fill("");
    const rowBg = new Array(PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS).fill(null);

    const stock = stockMap.get(key064) || {
      all: createProductListSizeObject_(),
      wh: createProductListSizeObject_(),
      allTotal: 0,
      whTotal: 0
    };

    // A〜AAなど、商品リストの6行目にある項目IDを見てVaMASTERから持ってくる
    Object.keys(tgtColMap).forEach(id => {
      const tgtCol = tgtColMap[id];

      // 固定マトリックス・管理列には通常ID転記しない
      if (tgtCol >= PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_START_COL) return;

      // 15_卸価格(¥) は式や手入力を残す想定なので、VaMASTERからは上書きしない
      if (id === "15") return;

      // 067_参照元商品URL：066の参照元コードから、VaMASTERの03を取る
      if (id === "067") {
        const refCode = vaColMap["066"]
          ? String(vaRow[vaColMap["066"] - 1] || "").trim()
          : "";

        const refInfo = refCode ? vaBy064.get(refCode) : null;

        if (refInfo && vaColMap["03"]) {
          rowData[tgtCol - 1] = refInfo.row[vaColMap["03"] - 1];
        }

        return;
      }

      // 068_参照元写真集：066の参照元コードから、VaMASTERの22を取る
      if (id === "068") {
        const refCode = vaColMap["066"]
          ? String(vaRow[vaColMap["066"] - 1] || "").trim()
          : "";

        const refInfo = refCode ? vaBy064.get(refCode) : null;

        if (refInfo && vaColMap["22"]) {
          rowData[tgtCol - 1] = refInfo.row[vaColMap["22"] - 1];

          // スマートチップ維持用
          copyTasks.push({
            srcSheet: vaSheet,
            srcRow: refInfo.rowNumber,
            srcCol: vaColMap["22"],
            dstRowOffset: outputRows.length,
            dstCol: tgtCol
          });
        }

        return;
      }

      // 写真表示
      if (id === "04" && vaColMap["05"]) {
        const photoUrl = vaRow[vaColMap["05"] - 1];
        if (photoUrl) {
          rowData[tgtCol - 1] = `=IMAGE("${getProductListDirectImageUrl_(photoUrl)}")`;
        }
        return;
      }

      // スマートチップ・リンク維持したい項目
      if (PRODUCT_LIST_CONFIG.COPY_TO_IDS.includes(id) && vaColMap[id]) {
        copyTasks.push({
          srcSheet: vaSheet,
          srcRow: PRODUCT_LIST_CONFIG.DATA_START_ROW + vaIndex,
          srcCol: vaColMap[id],
          dstRowOffset: outputRows.length,
          dstCol: tgtCol
        });

        return;
      }

      if (vaColMap[id]) {
        rowData[tgtCol - 1] = vaRow[vaColMap[id] - 1];
      }
    });

    // 既存手入力の復元
    if (backupMap.has(key064)) {
      const saved = backupMap.get(key064);

      Object.keys(saved).forEach(colText => {
        const col = Number(colText);
        rowData[col - 1] = saved[col];
      });
    }

    // 全在庫 AI:AO
    PRODUCT_LIST_CONFIG.SIZE_ORDER.forEach((size, idx) => {
      rowData[PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_START_COL - 1 + idx] =
        stock.all[size] || "";
    });
    rowData[PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_SUM_COL - 1] =
      stock.allTotal || "";

    // 倉庫在庫 AP:AV
    PRODUCT_LIST_CONFIG.SIZE_ORDER.forEach((size, idx) => {
      rowData[PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_START_COL - 1 + idx] =
        stock.wh[size] || "";
    });
    rowData[PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_SUM_COL - 1] =
      stock.whTotal || "";

    // 表示区分・並び順
    const sortNo = stock.allTotal > 0 ? 1 : 2;
    const displayType = stock.allTotal > 0 ? "在庫あり" : "在庫なし";

    rowData[PRODUCT_LIST_CONFIG.TARGET.SORT_COL - 1] = sortNo;
    rowData[PRODUCT_LIST_CONFIG.TARGET.DISPLAY_TYPE_COL - 1] = displayType;

    // 販売ステータスが空なら初期値
    if (!rowData[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL - 1]) {
      rowData[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL - 1] =
        stock.allTotal > 0 ? "販売確認中" : "在庫なし";
    }

    // 背景色
    for (
      let c = PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_START_COL;
      c <= PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_SUM_COL;
      c++
    ) {
      rowBg[c - 1] = "#eaf4ff";
    }

    for (
      let c = PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_START_COL;
      c <= PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_SUM_COL;
      c++
    ) {
      rowBg[c - 1] = "#f3f3f3";
    }

    rowBg[PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_SUM_COL - 1] = "#fff2cc";
    rowBg[PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_SUM_COL - 1] = "#fff2cc";

    rowBg[PRODUCT_LIST_CONFIG.TARGET.SORT_COL - 1] = "#eeeeee";
    rowBg[PRODUCT_LIST_CONFIG.TARGET.DISPLAY_TYPE_COL - 1] = "#eeeeee";

    outputRows.push(rowData);
    outputBgs.push(rowBg);
  });

  if (outputRows.length === 0) {
    Browser.msgBox("商品リストに展開できる商品がありませんでした。");
    return;
  }

  // ==========================================
  // Step 6: 在庫あり → 在庫なし の順に並び替え
  // ==========================================

  const combined = outputRows.map((row, i) => ({
    row,
    bg: outputBgs[i]
  }));

  combined.sort((a, b) => {
    const sortA = Number(a.row[PRODUCT_LIST_CONFIG.TARGET.SORT_COL - 1] || 9);
    const sortB = Number(b.row[PRODUCT_LIST_CONFIG.TARGET.SORT_COL - 1] || 9);

    if (sortA !== sortB) return sortA - sortB;

    const brandA = String(a.row[2 - 1] || "");
    const brandB = String(b.row[2 - 1] || "");
    if (brandA !== brandB) return brandA.localeCompare(brandB, "ja");

    const codeA = String(a.row[tgtColMap["064"] - 1] || "");
    const codeB = String(b.row[tgtColMap["064"] - 1] || "");
    return codeA.localeCompare(codeB, "ja");
  });

  const sortedRows = combined.map(x => x.row);
  const sortedBgs = combined.map(x => x.bg);

  // copyTasksは元のoutputRows順なので、ソート後の行位置を作り直すためのMap
  const rowIndexBy064 = new Map();
  sortedRows.forEach((row, idx) => {
    const key064 = String(row[tgtColMap["064"] - 1] || "").trim();
    if (key064) {
      rowIndexBy064.set(key064, idx);
    }
  });

  // ==========================================
  // Step 7: シートへ書き込み
  // ==========================================

  const neededMaxRows = PRODUCT_LIST_CONFIG.DATA_START_ROW + sortedRows.length - 1;
  const currentMaxRows = targetSheet.getMaxRows();

  if (neededMaxRows > currentMaxRows) {
    targetSheet.insertRowsAfter(currentMaxRows, neededMaxRows - currentMaxRows);
  }

  targetSheet.getRange(
    PRODUCT_LIST_CONFIG.DATA_START_ROW,
    1,
    sortedRows.length,
    PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS
  ).setValues(sortedRows);

  targetSheet.getRange(
    PRODUCT_LIST_CONFIG.DATA_START_ROW,
    1,
    sortedBgs.length,
    PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS
  ).setBackgrounds(sortedBgs);

  // ==========================================
  // Step 8: スマートチップ・リンク系 copyTo
  // ==========================================

  copyTasks.forEach(task => {
    const sourceRowData = outputRows[task.dstRowOffset];
    const key064 = String(sourceRowData[tgtColMap["064"] - 1] || "").trim();
    const sortedIndex = rowIndexBy064.get(key064);

    if (sortedIndex === undefined) return;

    const dstRow = PRODUCT_LIST_CONFIG.DATA_START_ROW + sortedIndex;

    task.srcSheet
      .getRange(task.srcRow, task.srcCol)
      .copyTo(targetSheet.getRange(dstRow, task.dstCol));

    targetSheet.getRange(dstRow, task.dstCol).setBackground(null);
  });

  // ==========================================
  // Step 9: 完了
  // ==========================================

  try {
    SpreadsheetApp.getUi().alert(
      "商品リストシート構築完了\n\n" +
      "展開商品数：" + sortedRows.length + "件\n" +
      "表示対象：一度でも入庫・受け入れされた商品"
    );
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "商品リストシート構築完了",
      "完了",
      10
    );
  }
}


/*******************************************************
 * ヘルパー
 *******************************************************/

function getProductListColMap_(sheet) {
  const headers = sheet.getRange(
    PRODUCT_LIST_CONFIG.HEADER_ROW,
    1,
    1,
    Math.max(sheet.getLastColumn(), 1)
  ).getValues()[0];

  const map = {};

  headers.forEach((header, idx) => {
    const text = normalizeProductListId_(header);
    const match = text.match(/^(\d{2,4})_/);

    if (match) {
      map[match[1]] = idx + 1;
    }
  });

  return map;
}

function normalizeProductListId_(value) {
  return String(value || "")
    .trim()
    .replace(/[０-９]/g, s =>
      String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    )
    .replace(/＿/g, "_");
}

function normalizeProductListSize_(value) {
  const size = String(value || "")
    .trim()
    .toUpperCase();

  if (
    size === "FREE" ||
    size === "FREES" ||
    size === "FREE SIZE" ||
    size === "フリー"
  ) {
    return "F";
  }

  return size;
}

function createProductListSizeObject_() {
  return {
    XS: 0,
    S: 0,
    M: 0,
    L: 0,
    XL: 0,
    F: 0
  };
}

function toProductListNumber_(value) {
  if (value === "" || value === null || value === undefined) return 0;

  const num = Number(value);

  return isNaN(num) ? 0 : num;
}

function getProductListDirectImageUrl_(url) {
  if (!url) return "";

  const match = String(url).match(/(?:id=|d\/)([\w-]+)/);

  if (match) {
    return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  }

  return url;
}