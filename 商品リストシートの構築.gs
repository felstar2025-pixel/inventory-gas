/*******************************************************
 * 商品リストシート構築 V1.3.5【15式実貼付・AB〜AH手入力保持版】
 *
 * 目的：
 * - 一度でも入庫・受け入れされた商品だけを表示
 * - VaMASTERの商品情報を商品リストへ展開
 * - 064合鍵で商品リスト側の手入力データを退避・復元
 * - 在庫マトリックスは「数値貼り付け」ではなく、倉庫マトリックス同様にVLOOKUP関数でSKUへ追従
 * - 存在しないサイズはグレーアウト
 * - VN → CN → その他、型番、BC優先、バリエーション、064順で並べる
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

  TARGET: {
    TOTAL_COLS: 55, // BC列まで（AB〜AHは手入力保持、AI以降は固定列）

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

  // 商品リスト側で手入力・編集される可能性が高い項目ID
  // ※15_卸価格(¥) はGASが7行目にBYROW式を貼るため、保存対象にしない。
  MANUAL_KEEP_IDS: [
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

  // VaMASTERからcopyToで維持したい項目ID
  // スマートチップ維持の対象は写真集だけ
  // 22: この商品の写真集
  // 068: 参照元写真集は特殊処理で、参照元行の22をcopyToする
  COPY_TO_IDS: ["22"],

  // 商品リスト側で自由入力欄として残す列位置
  // AB〜AH = 28〜34
  // 064をキーにして再構築後も復元する
  EXTRA_KEEP_COLS: [28, 29, 30, 31, 32, 33, 34],

  // 15_卸価格(¥) のBYROW式で使う為替レートセル
  // 必要ならここだけ変えれば、貼られる式も変わる
  EXCHANGE_RATE_CELLS: {
    VN: "$N$2",
    CN: "$N$3"
  },

  // 固定列の見出し名。6行目に入っていなくても、コード側でこの列を固定利用する
  FIXED_HEADERS: {
    ALL: ["全在庫_XS", "全在庫_S", "全在庫_M", "全在庫_L", "全在庫_XL", "全在庫_F", "2001_全在庫合計"],
    WH:  ["倉庫_XS", "倉庫_S", "倉庫_M", "倉庫_L", "倉庫_XL", "倉庫_F", "52_[倉庫] 現在庫"],
    CONTROL: ["4107_並び順", "4106_表示区分", "4040_販売ステータス"]
  }
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
    "商品リストシートを再構築します。\n" +
    "在庫マトリックスはSKU参照の関数で作り直します。\n" +
    "販売価格・説明文・ステータスなどは064コードをキーに復元します。\n\n" +
    "実行しますか？",
    Browser.Buttons.YES_NO
  );

  if (confirm !== "yes") return;

  const report = createProductListReport_();

  const vaColMap = getProductListColMap_(vaSheet);
  const skuColMap = getProductListColMap_(skuSheet);
  const tgtColMap = getProductListColMap_(targetSheet);

  if (!vaColMap["064"]) return Browser.msgBox("VaMASTERに 064_ 商品コード列が見つかりません。");
  if (!skuColMap["064"]) return Browser.msgBox("SKUシートに 064_ 商品コード列が見つかりません。AX列などに064生成列を用意してください。");
  if (!skuColMap["09"]) return Browser.msgBox("SKUシートに 09_ サイズ列が見つかりません。");
  if (!tgtColMap["064"]) return Browser.msgBox("商品リストシートに 064_ 商品コード列が見つかりません。");

  // 固定列の見出しを補正
  // 不足している項目IDは処理を止めず、最後にレポート表示する
  collectProductListMissingIdReport_(report, vaColMap, skuColMap, tgtColMap);

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

      // まず明示した手入力IDを保存
      PRODUCT_LIST_CONFIG.MANUAL_KEEP_IDS.forEach(id => {
        if (!tgtColMap[id]) return;
        const col = tgtColMap[id];
        saved[col] = row[col - 1];
      });

      // 固定列側の販売ステータスも保存
      saved[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL] =
        row[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL - 1];

      // AA〜AGの自由入力欄も列位置で保存
      PRODUCT_LIST_CONFIG.EXTRA_KEEP_COLS.forEach(col => {
        if (col <= row.length) saved[col] = row[col - 1];
      });

      backupMap.set(key064, saved);
    });
  }

  // ==========================================
  // Step 2: SKUから「一度でも扱った064」を判定
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

  const intakeIds = ["50", "54", "58", "71", "75"];
  const allStockId = "2001";
  const whStockId = "52";

  const handled064Set = new Set();
  const initialStockMap = new Map();

  skuValues.forEach(row => {
    const key064 = String(row[skuColMap["064"] - 1] || "").trim();
    if (!key064) return;

    const size = normalizeProductListSize_(row[skuColMap["09"] - 1]);
    if (!PRODUCT_LIST_CONFIG.SIZE_ORDER.includes(size)) return;

    const intakeTotal = intakeIds.reduce((sum, id) => {
      if (!skuColMap[id]) return sum;
      return sum + toProductListNumber_(row[skuColMap[id] - 1]);
    }, 0);

    if (intakeTotal > 0) handled064Set.add(key064);

    if (!initialStockMap.has(key064)) {
      initialStockMap.set(key064, {
        allTotal: 0,
        whTotal: 0
      });
    }

    const stockInfo = initialStockMap.get(key064);

    let allQty = 0;
    if (skuColMap[allStockId]) {
      allQty = toProductListNumber_(row[skuColMap[allStockId] - 1]);
    } else {
      ["52", "56", "60", "73", "76"].forEach(id => {
        if (skuColMap[id]) allQty += toProductListNumber_(row[skuColMap[id] - 1]);
      });
    }

    const whQty = skuColMap[whStockId]
      ? toProductListNumber_(row[skuColMap[whStockId] - 1])
      : 0;

    stockInfo.allTotal += allQty;
    stockInfo.whTotal += whQty;
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

  const vaBy064 = new Map();
  const sizeExistMap = new Map();

  vaValues.forEach((row, idx) => {
    const key064 = String(row[vaColMap["064"] - 1] || "").trim();
    if (!key064) return;

    vaBy064.set(key064, {
      row,
      rowNumber: PRODUCT_LIST_CONFIG.DATA_START_ROW + idx
    });

    const sizeText = vaColMap["09"] ? row[vaColMap["09"] - 1] : "";
    sizeExistMap.set(key064, getProductListAvailableSizes_(sizeText));
  });

  // ==========================================
  // Step 4: 商品リストの初期化 倉庫方式
  // ==========================================
  if (targetLastRow >= PRODUCT_LIST_CONFIG.DATA_START_ROW) {
    clearProductListContentAndProtections_(targetSheet.getRange(
      PRODUCT_LIST_CONFIG.DATA_START_ROW,
      1,
      targetLastRow - PRODUCT_LIST_CONFIG.DATA_START_ROW + 1,
      PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS
    ));
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
    if (!handled064Set.has(key064)) return;

    const rowData = new Array(PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS).fill("");
    const rowBg = new Array(PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS).fill(null);

    const actualSizes = sizeExistMap.get(key064) || new Set();
    const stockInfo = initialStockMap.get(key064) || { allTotal: 0, whTotal: 0 };

    // 商品情報側：商品リスト6行目の項目IDに合わせてVaMASTERから持ってくる
    Object.keys(tgtColMap).forEach(id => {
      const tgtCol = tgtColMap[id];
      if (tgtCol >= PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_START_COL) return;

        // 15_卸価格(¥) は再構築後にBYROW式を貼るため、VaMASTERからは転記しない。
      if (id === "15") return;

      // 067_参照元商品URL：066の参照元コードからVaMASTERの03を取る
      if (id === "067") {
        const refCode = vaColMap["066"]
          ? String(vaRow[vaColMap["066"] - 1] || "").trim()
          : "";
        const refInfo = refCode ? vaBy064.get(refCode) : null;
        if (refInfo && vaColMap["03"]) rowData[tgtCol - 1] = refInfo.row[vaColMap["03"] - 1];
        return;
      }

      // 068_参照元写真集：066の参照元コードからVaMASTERの22を取る
      if (id === "068") {
        const refCode = vaColMap["066"]
          ? String(vaRow[vaColMap["066"] - 1] || "").trim()
          : "";
        const refInfo = refCode ? vaBy064.get(refCode) : null;
        if (refInfo && vaColMap["22"]) {
          rowData[tgtCol - 1] = refInfo.row[vaColMap["22"] - 1];
          copyTasks.push({
            srcSheet: vaSheet,
            srcRow: refInfo.rowNumber,
            srcCol: vaColMap["22"],
            dstRowOffset: outputRows.length,
            dstCol: tgtCol,
            key064
          });
        }
        return;
      }

      // 写真表示
      if (id === "04" && vaColMap["05"]) {
        const photoUrl = vaRow[vaColMap["05"] - 1];
        if (photoUrl) rowData[tgtCol - 1] = `=IMAGE("${getProductListDirectImageUrl_(photoUrl)}")`;
        return;
      }

      // サイズはVaMASTERの09から存在サイズだけを整形
      if (id === "09") {
        rowData[tgtCol - 1] = PRODUCT_LIST_CONFIG.SIZE_ORDER.filter(s => actualSizes.has(s)).join(", ");
        return;
      }

      // copyToしたい項目。ただし手入力バックアップがある列は後で上書きしない
      if (PRODUCT_LIST_CONFIG.COPY_TO_IDS.includes(id) && vaColMap[id]) {
        copyTasks.push({
          srcSheet: vaSheet,
          srcRow: PRODUCT_LIST_CONFIG.DATA_START_ROW + vaIndex,
          srcCol: vaColMap[id],
          dstRowOffset: outputRows.length,
          dstCol: tgtCol,
          key064,
          id
        });
        return;
      }

      if (vaColMap[id]) rowData[tgtCol - 1] = vaRow[vaColMap[id] - 1];
    });

    // 手入力データの復元
    if (backupMap.has(key064)) {
      const saved = backupMap.get(key064);
      Object.keys(saved).forEach(colText => {
        const col = Number(colText);
        rowData[col - 1] = saved[col];
      });
    }

    // 表示区分・並び順の初期値。ARRAYFORMULAは貼らず、再構築時点の値として保持
    const sortNo = stockInfo.allTotal > 0 ? 1 : 2;
    rowData[PRODUCT_LIST_CONFIG.TARGET.SORT_COL - 1] = sortNo;
    rowData[PRODUCT_LIST_CONFIG.TARGET.DISPLAY_TYPE_COL - 1] = stockInfo.allTotal > 0 ? "在庫あり" : "在庫なし";

    if (!rowData[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL - 1]) {
      rowData[PRODUCT_LIST_CONFIG.TARGET.SALES_STATUS_COL - 1] = stockInfo.allTotal > 0 ? "販売確認中" : "在庫なし";
    }

    // 倉庫方式の色塗り：存在しないサイズはグレー、合計列は黄色
    applyProductListMatrixBackground_(rowBg, actualSizes);
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
  // Step 6: 並び替え 倉庫・総合マトリックス寄り
  // ==========================================
  const combined = outputRows.map((row, i) => ({ row, bg: outputBgs[i] }));

  combined.sort((a, b) => {
    // 1. 在庫あり / 在庫なし
    const sortA = Number(a.row[PRODUCT_LIST_CONFIG.TARGET.SORT_COL - 1] || 9);
    const sortB = Number(b.row[PRODUCT_LIST_CONFIG.TARGET.SORT_COL - 1] || 9);
    if (sortA !== sortB) return sortA - sortB;

    const codeA = String(a.row[tgtColMap["064"] - 1] || "");
    const codeB = String(b.row[tgtColMap["064"] - 1] || "");

    const vaA = vaBy064.get(codeA);
    const vaB = vaBy064.get(codeB);
    const vaRowA = vaA ? vaA.row : [];
    const vaRowB = vaB ? vaB.row : [];

    // 2. 国順 VN -> CN -> その他
    const countryA = getProductListCountrySortNoFromVa_(vaRowA, vaColMap);
    const countryB = getProductListCountrySortNoFromVa_(vaRowB, vaColMap);
    if (countryA !== countryB) return countryA - countryB;

    // 3. 型番コード 06
    const modelA = vaColMap["06"] ? String(vaRowA[vaColMap["06"] - 1] || "") : getProductListCodePart_(codeA, 0);
    const modelB = vaColMap["06"] ? String(vaRowB[vaColMap["06"] - 1] || "") : getProductListCodePart_(codeB, 0);
    if (modelA !== modelB) return modelA.localeCompare(modelB, "ja");

    // 4. サプライヤー BC優先
    const supplierA = getProductListSupplierCodeFrom064_(codeA);
    const supplierB = getProductListSupplierCodeFrom064_(codeB);
    const supSortA = getProductListSupplierSortNo_(supplierA);
    const supSortB = getProductListSupplierSortNo_(supplierB);
    if (supSortA !== supSortB) return supSortA - supSortB;
    if (supplierA !== supplierB) return supplierA.localeCompare(supplierB, "ja");

    // 5. バリエーション順
    const variationA = getProductListVariationCodeFrom064_(codeA);
    const variationB = getProductListVariationCodeFrom064_(codeB);
    const varSortA = getProductListVariationSortNo_(variationA);
    const varSortB = getProductListVariationSortNo_(variationB);
    if (varSortA !== varSortB) return varSortA - varSortB;
    if (variationA !== variationB) return variationA.localeCompare(variationB, "ja");

    // 6. 最後に064
    return codeA.localeCompare(codeB, "ja");
  });

  const sortedRows = combined.map(x => x.row);
  const sortedBgs = combined.map(x => x.bg);

  const rowIndexBy064 = new Map();
  sortedRows.forEach((row, idx) => {
    const key064 = String(row[tgtColMap["064"] - 1] || "").trim();
    if (key064) rowIndexBy064.set(key064, idx);
  });

  // ==========================================
  // Step 7: 書き込み
  // ==========================================
  const neededMaxRows = PRODUCT_LIST_CONFIG.DATA_START_ROW + sortedRows.length - 1;
  const currentMaxRows = targetSheet.getMaxRows();
  if (neededMaxRows > currentMaxRows) {
    targetSheet.insertRowsAfter(currentMaxRows, neededMaxRows - currentMaxRows);
  }

  // 15_卸価格(¥) 列はこのあとBYROW式を貼るため、
  // setValues / setBackgrounds の対象から外す。
  writeProductListRowsExcludingColumns_(
    targetSheet,
    PRODUCT_LIST_CONFIG.DATA_START_ROW,
    sortedRows,
    sortedBgs,
    [tgtColMap["15"]].filter(Boolean)
  );

  // ==========================================
  // Step 8: スマートチップ・リンク系 copyTo
  // ==========================================
  copyTasks.forEach(task => {
    const sourceRowData = outputRows[task.dstRowOffset];
    const key064 = String(sourceRowData[tgtColMap["064"] - 1] || "").trim();
    const sortedIndex = rowIndexBy064.get(key064);
    if (sortedIndex === undefined) return;

    // copyTo対象は写真集系だけ。手入力欄には使わない。
    const dstRow = PRODUCT_LIST_CONFIG.DATA_START_ROW + sortedIndex;
    task.srcSheet.getRange(task.srcRow, task.srcCol).copyTo(targetSheet.getRange(dstRow, task.dstCol));
    targetSheet.getRange(dstRow, task.dstCol).setBackground(null);
  });

  // ==========================================
  // Step 9: 倉庫方式の関数・保護を貼る
  // ==========================================
  const fStart = PRODUCT_LIST_CONFIG.DATA_START_ROW;
  const finalLastRow = neededMaxRows;
  const col064Str = getProductListColumnLetter_(tgtColMap["064"]);

  // ① 15_卸価格(¥) の円換算式
  //    13_購入通貨 と 14_卸価格(NT) を、項目IDで見つけた列から参照する。
  //    CHOOSEで2列だけを組み直すので、13列と14列の間に別列があってもズレない。
  const colJpy = tgtColMap["15"];
  if (colJpy && tgtColMap["13"] && tgtColMap["14"]) {
    const countryColLetter = getProductListColumnLetter_(tgtColMap["13"]);
    const priceColLetter = getProductListColumnLetter_(tgtColMap["14"]);

    const formula =
  `=BYROW(${getProductListColumnLetter_(tgtColMap["13"])}${fStart}:${getProductListColumnLetter_(tgtColMap["14"])}${finalLastRow}, ` +
  `LAMBDA(row, IF(INDEX(row, 1, 2)="", "", ` +
  `IF(INDEX(row, 1, 1)="VN", INDEX(row, 1, 2) * ${PRODUCT_LIST_CONFIG.EXCHANGE_RATE_CELLS.VN}, ` +
  `IF(INDEX(row, 1, 1)="CN", INDEX(row, 1, 2) * ${PRODUCT_LIST_CONFIG.EXCHANGE_RATE_CELLS.CN}, "")))))`;
  
    targetSheet.getRange(fStart, colJpy).setFormula(formula);
  } else {
    const missing = [];
    if (!tgtColMap["13"]) missing.push("13_購入通貨");
    if (!tgtColMap["14"]) missing.push("14_卸価格(NT)");
    if (!tgtColMap["15"]) missing.push("15_卸価格(¥)");

    report.notes.push(
      "15_卸価格(¥) の円換算式は貼っていません。商品リストに " +
      missing.join(" / ") +
      " が見つかりませんでした。"
    );
  }

  // ② 全在庫・倉庫在庫のサイズ別VLOOKUP関数
  const skuLookupLastCol = Math.max(skuSheet.getLastColumn(), skuColMap[allStockId] || 1, skuColMap[whStockId] || 1);
  const skuLookupRange = `'${PRODUCT_LIST_CONFIG.SOURCE_SKU}'!$A:$${getProductListColumnLetter_(skuLookupLastCol)}`;

  if (skuColMap[allStockId]) {
    const allFormulas = [];
    for (let r = fStart; r <= finalLastRow; r++) {
      const rowF = PRODUCT_LIST_CONFIG.SIZE_ORDER.map(size => {
        return `=IF($${col064Str}${r}="", "", IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", ${skuLookupRange}, ${skuColMap[allStockId]}, FALSE), ""))`;
      });
      allFormulas.push(rowF);
    }
    targetSheet.getRange(fStart, PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_START_COL, allFormulas.length, 6).setFormulas(allFormulas);
  } else {
    report.warnings.push("SKUシートに 2001_全在庫合計 が見つからないため、全在庫マトリックスのVLOOKUP式は貼っていません。");
  }

  if (skuColMap[whStockId]) {
    const whFormulas = [];
    for (let r = fStart; r <= finalLastRow; r++) {
      const rowF = PRODUCT_LIST_CONFIG.SIZE_ORDER.map(size => {
        return `=IF($${col064Str}${r}="", "", IFERROR(VLOOKUP($${col064Str}${r} & "-${size}", ${skuLookupRange}, ${skuColMap[whStockId]}, FALSE), ""))`;
      });
      whFormulas.push(rowF);
    }
    targetSheet.getRange(fStart, PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_START_COL, whFormulas.length, 6).setFormulas(whFormulas);
  } else {
    report.warnings.push("SKUシートに 52_[倉庫] 現在庫 が見つからないため、倉庫在庫マトリックスのVLOOKUP式は貼っていません。");
  }

  // ③ 合計列はBYROWでリアルタイム追従
  targetSheet.getRange(fStart, PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_SUM_COL)
    .setFormula(`=BYROW(AI${fStart}:AN${finalLastRow}, LAMBDA(row, IF(COUNTA(row)=0, "", SUM(row))))`);

  targetSheet.getRange(fStart, PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_SUM_COL)
    .setFormula(`=BYROW(AP${fStart}:AU${finalLastRow}, LAMBDA(row, IF(COUNTA(row)=0, "", SUM(row))))`);

  // ④ 並び順・表示区分は、再構築時点の値として保持する。
  // ※ここにARRAYFORMULAを貼ると、すでにsetValuesで入れたAW/AXの値と衝突して
  //   「配列結果を展開できません」エラーになる。
  // ※行の並び替えはGAS再構築時に行うため、AW/AXだけリアルタイム関数化しない。

  // ⑤ 関数列・在庫列に警告保護
  const protectRanges = [
    targetSheet.getRange(`AI${fStart}:AO${finalLastRow}`),
    targetSheet.getRange(`AP${fStart}:AV${finalLastRow}`),
    targetSheet.getRange(`AW${fStart}:AX${finalLastRow}`)
  ];

  if (tgtColMap["15"] && tgtColMap["13"] && tgtColMap["14"]) {
    protectRanges.push(targetSheet.getRange(fStart, tgtColMap["15"], finalLastRow - fStart + 1, 1));
  }


  protectRanges.forEach(rng => rng.protect().setWarningOnly(true));

  const resultMessage = buildProductListResultMessage_(sortedRows.length, report);

  try {
    SpreadsheetApp.getUi().alert(resultMessage);
  } catch (e) {
    ss.toast("商品リストシート構築完了", "完了", 10);
  }
}

/*******************************************************
 * ヘルパー
 *******************************************************/

function writeProductListRowsExcludingColumns_(sheet, startRow, rows, backgrounds, excludeCols) {
  const excludeSet = new Set((excludeCols || []).filter(Boolean));
  const totalCols = PRODUCT_LIST_CONFIG.TARGET.TOTAL_COLS;
  const rowCount = rows.length;

  let startCol = 1;

  while (startCol <= totalCols) {
    while (startCol <= totalCols && excludeSet.has(startCol)) startCol++;
    if (startCol > totalCols) break;

    let endCol = startCol;
    while (endCol + 1 <= totalCols && !excludeSet.has(endCol + 1)) endCol++;

    const width = endCol - startCol + 1;
    const valuesPart = rows.map(row => row.slice(startCol - 1, endCol));
    const bgsPart = backgrounds.map(row => row.slice(startCol - 1, endCol));

    sheet.getRange(startRow, startCol, rowCount, width).setValues(valuesPart);
    sheet.getRange(startRow, startCol, rowCount, width).setBackgrounds(bgsPart);

    startCol = endCol + 1;
  }
}

function createProductListReport_() {
  return {
    warnings: [],
    notes: []
  };
}

function collectProductListMissingIdReport_(report, vaColMap, skuColMap, tgtColMap) {
  // 商品リスト側に列がある場合だけ確認する。
  // 手入力専用IDはVaMASTERに無くて正常なので、ここでは確認しない。
  const vaRequiredByTarget = {
    "064": "064_商品コード[Patt+Vari+Sup]",
    "066": "066_参照元商品ｺｰﾄﾞ",
    "02": "02_店舗/BRAND",
    "01": "01_サプライヤー",
    "05": "05_写真URL",
    "03": "03_商品URL",
    "10": "10_Variation（NT）",
    "12": "12_Variation（JP）",
    "09": "09_Size",
    "13": "13_購入通貨",
    "14": "14_卸価格(NT)",
    "22": "22_写真集"
  };

  Object.keys(vaRequiredByTarget).forEach(id => {
    if (tgtColMap[id] && !vaColMap[id]) {
      report.warnings.push("VaMASTERに " + vaRequiredByTarget[id] + " が見つかりませんでした。");
    }
  });

  if (tgtColMap["04"] && !vaColMap["05"]) {
    report.warnings.push("商品リストに 04_写真 がありますが、VaMASTERに 05_写真URL がないため、画像式を作れません。");
  }

  if (tgtColMap["067"]) {
    if (!vaColMap["066"]) report.warnings.push("067_参照元商品URL を作るための 066_参照元商品ｺｰﾄﾞ がVaMASTERに見つかりませんでした。");
    if (!vaColMap["03"]) report.warnings.push("067_参照元商品URL を作るための 03_商品URL がVaMASTERに見つかりませんでした。");
  }

  if (tgtColMap["068"]) {
    if (!vaColMap["066"]) report.warnings.push("068_参照元写真集 を作るための 066_参照元商品ｺｰﾄﾞ がVaMASTERに見つかりませんでした。");
    if (!vaColMap["22"]) report.warnings.push("068_参照元写真集 を作るための 22_写真集 がVaMASTERに見つかりませんでした。");
  }

  const skuRequired = {
    "09": "09_Size",
    "064": "064_商品コード[Patt+Vari+Sup]",
    "2001": "2001_全在庫合計",
    "52": "52_[倉庫] 現在庫"
  };

  Object.keys(skuRequired).forEach(id => {
    if (!skuColMap[id]) {
      report.warnings.push("SKUシートに " + skuRequired[id] + " が見つかりませんでした。");
    }
  });
}

function buildProductListResultMessage_(rowCount, report) {
  const lines = [
    "商品リストシート構築完了",
    "",
    "展開商品数：" + rowCount + "件",
    "在庫マトリックス：SKU参照関数で作成",
    "手入力データ：指定ID＋AB〜AHを064コードで復元",
    "copyTo対象：22_写真集 / 068_参照元写真集",
    "15_卸価格(¥)：7行目にBYROW式を貼付",
    ""
  ];

  if (report.warnings.length || report.notes.length) {
    lines.push("【確認事項】");
    report.warnings.forEach(msg => lines.push("・" + msg));
    report.notes.forEach(msg => lines.push("・" + msg));
    lines.push("");
    lines.push("※処理は完了しています。");
  } else {
    lines.push("確認事項：なし");
  }

  return lines.join("\n");
}

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
    if (match) map[match[1]] = idx + 1;
  });
  return map;
}

function normalizeProductListId_(value) {
  return String(value || "")
    .trim()
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/＿/g, "_");
}

function normalizeProductListSize_(value) {
  const size = String(value || "").trim().toUpperCase();
  if (["FREE", "FREES", "FREE SIZE", "フリー"].includes(size)) return "F";
  return size;
}

function getProductListAvailableSizes_(sizeText) {
  const result = new Set();
  const text = String(sizeText || "").trim();

  if (!text) return result;

  text.split(/[,、\n]/)
    .map(size => normalizeProductListSize_(size))
    .filter(Boolean)
    .forEach(size => {
      if (PRODUCT_LIST_CONFIG.SIZE_ORDER.includes(size)) result.add(size);
    });

  return result;
}

function toProductListNumber_(value) {
  if (value === "" || value === null || value === undefined) return 0;
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

function getProductListDirectImageUrl_(url) {
  if (!url) return "";
  const match = String(url).match(/(?:id=|d\/)([\w-]+)/);
  if (match) return `https://drive.google.com/uc?export=download&id=${match[1]}`;
  return url;
}

function getProductListColumnLetter_(column) {
  let temp;
  let letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function clearProductListContentAndProtections_(range) {
  range.clearContent();
  range.setBackground(null);

  const sheet = range.getSheet();
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => p.remove());

  // 商品リストでは条件付き書式は基本使わないが、倉庫方式に合わせて再構築時にクリアする
  sheet.clearConditionalFormatRules();
}

// ※v1.3.3では6行目の項目ID・見出しはGASで上書きしません。
// 固定ヘッダー自動書き込み処理は削除しました。

function applyProductListMatrixBackground_(rowBg, actualSizes) {
  // 全在庫 AI:AN、倉庫 AP:AU の通常色
  for (let c = PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_START_COL; c <= PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_SUM_COL; c++) {
    rowBg[c - 1] = "#f3f3f3";
  }

  for (let c = PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_START_COL; c <= PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_SUM_COL; c++) {
    rowBg[c - 1] = "#f3f3f3";
  }

  // 合計列
  rowBg[PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_SUM_COL - 1] = "#fff2cc";
  rowBg[PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_SUM_COL - 1] = "#fff2cc";

  // 存在しないサイズはグレー
  PRODUCT_LIST_CONFIG.SIZE_ORDER.forEach((size, idx) => {
    if (actualSizes.has(size)) return;
    rowBg[PRODUCT_LIST_CONFIG.TARGET.ALL_STOCK_START_COL - 1 + idx] = "#999999";
    rowBg[PRODUCT_LIST_CONFIG.TARGET.WH_STOCK_START_COL - 1 + idx] = "#999999";
  });
}

function getProductListCountrySortNoFromVa_(vaRow, vaColMap) {
  const text = vaColMap["13"]
    ? String(vaRow[vaColMap["13"] - 1] || "").trim().toUpperCase()
    : "";

  if (text.includes("VN") || text.includes("VND") || text.includes("VIETNAM") || text.includes("ベトナム")) return 1;
  if (text.includes("CN") || text.includes("CNY") || text.includes("RMB") || text.includes("CHINA") || text.includes("中国") || text.includes("元")) return 2;
  return 9;
}

function getProductListSupplierCodeFrom064_(code064) {
  const parts = String(code064 || "").split("-");
  return parts.length >= 3 ? parts[2] : "";
}

function getProductListVariationCodeFrom064_(code064) {
  const parts = String(code064 || "").split("-");
  return parts.length >= 2 ? parts[1] : "";
}

function getProductListCodePart_(code, index) {
  return String(code || "").split("-")[index] || "";
}

function getProductListSupplierSortNo_(supplierCode) {
  const code = String(supplierCode || "").trim().toUpperCase();
  if (code === "BC") return 1;
  return 2;
}

function getProductListVariationSortNo_(variationCode) {
  const code = String(variationCode || "").trim().toUpperCase();
  if (code === "N") return 1;
  if (code === "A") return 2;
  if (code === "B") return 3;
  if (code === "C") return 4;
  if (code === "D") return 5;
  const num = Number(code);
  if (!isNaN(num)) return 100 + num;
  return 999;
}
