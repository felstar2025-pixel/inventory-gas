/**
 * =========================================================================
 * サイドバー在庫閲覧GS.gs
 * Viewer版：sidebar_view_only.html 用
 *
 * 目的：
 * - sidebar_view_only.html はそのまま使う
 * - クリックされたシートからは「商品を特定するコード」だけ取る
 * - 表示する商品情報は VaMASTER から取得する
 * - 在庫情報は SKU から取得する
 *
 * クリック条件：
 * - 各シート6行目の項目IDが「04_」で始まる列をクリックした時だけ反応
 *
 * 主キー：
 * - 優先1：064_商品コード[Patt+Vari+Sup]
 * - 優先2：061_完全SKUコード から末尾サイズを除去して064化
 * - 優先3：06_商品コード（先頭7文字）でVaMASTERを予備検索
 * =========================================================================
 */


/**
 * sidebar_view_only.html から呼ばれる関数
 */
function getSelectedImageUrlViewer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const range = activeSheet.getActiveCell();

  const headerRow = 6;
  const row = range.getRow();
  const col = range.getColumn();

  if (row <= headerRow) return null;

  const currentHeaders = activeSheet
    .getRange(headerRow, 1, 1, activeSheet.getLastColumn())
    .getValues()[0];

  const clickedHeader = normalizeViewerHeader_(currentHeaders[col - 1]);

  // 04_写真列をクリックした時だけ反応
  if (!clickedHeader.startsWith("04_")) return null;

  const currentColMap = getViewerColMapFromHeaders_(currentHeaders);

  // クリックした行から商品特定コードを取得
  const clickedKey = getViewerClickedProductKey_(activeSheet, row, currentColMap);

  if (!clickedKey.rawKey) {
    return makeViewerEmptyResult_("-", "商品コードがありません");
  }

  const vaSheet = ss.getSheetByName("VaMASTER");
  if (!vaSheet) {
    return makeViewerEmptyResult_(clickedKey.displayCode, "VaMASTERシートが見つかりません");
  }

  const vaInfo = findViewerVaMasterRow_(vaSheet, clickedKey, headerRow);

  if (!vaInfo) {
    return makeViewerEmptyResult_(clickedKey.displayCode, "VaMASTERに該当商品がありません");
  }

  const vaRow = vaInfo.row;
  const vaColMap = vaInfo.colMap;

  // VaMASTERから表示情報を取得
  const code =
    getViewerValue_(vaRow, vaColMap, "064") ||
    getViewerValue_(vaRow, vaColMap, "06") ||
    clickedKey.displayCode;

  const primaryPhotoUrl =
    getViewerRichOrValue_(vaSheet, vaInfo.rowNumber, vaColMap["05"]) ||
    getViewerRichOrValue_(vaSheet, vaInfo.rowNumber, vaColMap["04"]) ||
    "";

  const siteUrl =
    getViewerRichOrValue_(vaSheet, vaInfo.rowNumber, vaColMap["03"]) || "";

  const nameEn =
    getViewerValue_(vaRow, vaColMap, "17") ||
    getViewerValue_(vaRow, vaColMap, "111") ||
    getViewerValue_(vaRow, vaColMap, "08") ||
    "";

  const nameJp =
    getViewerValue_(vaRow, vaColMap, "121") ||
    getViewerValue_(vaRow, vaColMap, "16") ||
    getViewerValue_(vaRow, vaColMap, "12") ||
    "";

  const priceVn = toViewerNumber_(getViewerValue_(vaRow, vaColMap, "14"));
  const priceJp = toViewerNumber_(getViewerValue_(vaRow, vaColMap, "15"));

  // 在庫情報はSKUから取得
  const key064 =
    getViewerValue_(vaRow, vaColMap, "064") ||
    clickedKey.key064 ||
    "";

  const fallbackBaseCode =
    getViewerValue_(vaRow, vaColMap, "06") ||
    clickedKey.baseCode ||
    "";

  const stockData = collectViewerStockData_(ss, key064, fallbackBaseCode, headerRow);

  const allUrls = String(primaryPhotoUrl || "")
    .split(/[,、\n\s]+/)
    .map(u => u.trim())
    .filter(u => u !== "")
    .map(getThumbnailUrlViewer);

  return {
    url: allUrls[0] || "",
    urls: allUrls,
    code: code,
    nameEn: nameEn,
    nameJp: nameJp,
    priceVn: priceVn,
    priceJp: priceJp,
    siteUrl: siteUrl,
    stockData: stockData
  };
}


/**
 * クリック行から商品特定キーを作る
 */
function getViewerClickedProductKey_(sheet, row, colMap) {
  // 優先1：064
  if (colMap["064"]) {
    const key064 = String(sheet.getRange(row, colMap["064"]).getValue() || "").trim();
    if (key064) {
      return {
        rawKey: key064,
        key064: key064,
        baseCode: key064.substring(0, 7).toUpperCase(),
        displayCode: key064,
        type: "064"
      };
    }
  }

  // 優先2：061 完全SKUコードから064を作る
  if (colMap["061"]) {
    const fullSku = String(sheet.getRange(row, colMap["061"]).getValue() || "").trim();
    const key064 = convertViewerSkuTo064_(fullSku);
    if (key064) {
      return {
        rawKey: fullSku,
        key064: key064,
        baseCode: key064.substring(0, 7).toUpperCase(),
        displayCode: key064,
        type: "061"
      };
    }
  }

  // 優先3：06 商品コード。曖昧なので予備
  if (colMap["06"]) {
    const rawCode = String(sheet.getRange(row, colMap["06"]).getValue() || "").trim();
    const baseCode = rawCode.substring(0, 7).toUpperCase();
    if (baseCode) {
      return {
        rawKey: rawCode,
        key064: "",
        baseCode: baseCode,
        displayCode: baseCode,
        type: "06"
      };
    }
  }

  return {
    rawKey: "",
    key064: "",
    baseCode: "",
    displayCode: "-",
    type: ""
  };
}


/**
 * VaMASTERから該当行を探す
 */
function findViewerVaMasterRow_(vaSheet, clickedKey, headerRow) {
  const lastRow = vaSheet.getLastRow();
  const lastCol = vaSheet.getLastColumn();

  if (lastRow <= headerRow) return null;

  const headers = vaSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const colMap = getViewerColMapFromHeaders_(headers);

  const values = vaSheet
    .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
    .getValues();

  // 064が取れているなら、VaMASTERの064で完全一致検索
  if (clickedKey.key064 && colMap["064"]) {
    for (let i = 0; i < values.length; i++) {
      const va064 = String(values[i][colMap["064"] - 1] || "").trim();
      if (va064 === clickedKey.key064) {
        return {
          row: values[i],
          rowNumber: headerRow + 1 + i,
          colMap: colMap
        };
      }
    }
  }

  // 064で見つからない場合だけ、06で予備検索
  // 06はバリエーション違い・サプライヤー違いが曖昧になるので最後の手段
  if (clickedKey.baseCode && colMap["06"]) {
    for (let i = 0; i < values.length; i++) {
      const va06 = String(values[i][colMap["06"] - 1] || "")
        .trim()
        .substring(0, 7)
        .toUpperCase();

      if (va06 === clickedKey.baseCode) {
        return {
          row: values[i],
          rowNumber: headerRow + 1 + i,
          colMap: colMap
        };
      }
    }
  }

  return null;
}


/**
 * SKUから在庫情報を集める
 */
function collectViewerStockData_(ss, key064, fallbackBaseCode, headerRow) {
  const skuSheet = ss.getSheetByName("SKU");
  if (!skuSheet) return [];

  const lastRow = skuSheet.getLastRow();
  const lastCol = skuSheet.getLastColumn();

  if (lastRow <= headerRow) return [];

  const headers = skuSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const colMap = getViewerColMapFromHeaders_(headers);

  const colFullSku = colMap["061"] || colMap["06"]; // 古いSKU構成の保険
  const colSku064 = colMap["064"];
  const colVar = colMap["12"] || colMap["11"] || colMap["10"];
  const colSize = colMap["09"];
  const colStock = colMap["52"]; // 倉庫現在庫

  if (!colFullSku && !colSku064) return [];

  const values = skuSheet
    .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
    .getValues();

  const result = [];

  values.forEach(row => {
    const sku064 = colSku064
      ? String(row[colSku064 - 1] || "").trim()
      : "";

    const fullSku = colFullSku
      ? String(row[colFullSku - 1] || "").trim()
      : "";

    const derived064 = sku064 || convertViewerSkuTo064_(fullSku);
    const baseCode = (derived064 || fullSku).substring(0, 7).toUpperCase();

    let matched = false;

    if (key064 && derived064 === key064) {
      matched = true;
    } else if (!key064 && fallbackBaseCode && baseCode === fallbackBaseCode.substring(0, 7).toUpperCase()) {
      matched = true;
    }

    if (!matched) return;

    result.push({
      variation: colVar ? String(row[colVar - 1] || "") : "基本",
      size: colSize ? String(row[colSize - 1] || "") : "-",
      stock: colStock ? Number(row[colStock - 1] || 0) : 0
    });
  });

  return result;
}


/**
 * 完全SKUコードから064を作る
 * 想定：BJ10001-N-BC-M → BJ10001-N-BC
 */
function convertViewerSkuTo064_(fullSku) {
  const text = String(fullSku || "").trim();
  if (!text) return "";

  return text.replace(/-(XS|S|M|L|XL|F|FREE|FREE SIZE|フリー)$/i, "");
}


/**
 * 6行目ヘッダーから項目IDマップを作る
 */
function getViewerColMapFromHeaders_(headers) {
  const map = {};

  headers.forEach((header, idx) => {
    const text = normalizeViewerHeader_(header);
    const match = text.match(/^(\d{2,4})_/);

    if (match) {
      map[match[1]] = idx + 1;
    }
  });

  return map;
}


function normalizeViewerHeader_(value) {
  return String(value || "")
    .trim()
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/＿/g, "_");
}


function getViewerValue_(row, colMap, id) {
  if (!colMap[id]) return "";
  return row[colMap[id] - 1];
}


function getViewerRichOrValue_(sheet, row, col) {
  if (!col) return "";

  const cell = sheet.getRange(row, col);

  try {
    const rich = cell.getRichTextValue();
    if (rich) {
      const link = rich.getLinkUrl();
      const text = rich.getText();

      if (link) return link;
      if (text) return text;
    }
  } catch (e) {
    // RichTextが取れない場合は通常値へ
  }

  return String(cell.getValue() || "");
}


function toViewerNumber_(value) {
  if (value === "" || value === null || value === undefined) return 0;

  const num = Number(value);

  return isNaN(num) ? 0 : num;
}


function makeViewerEmptyResult_(code, message) {
  return {
    url: "",
    urls: [],
    code: code || "-",
    nameEn: "-",
    nameJp: message || "-",
    priceVn: 0,
    priceJp: 0,
    siteUrl: "",
    stockData: []
  };
}


/**
 * 画像変換用：Google Drive URLをサイドバー表示用サムネイルURLに変換
 */
function getThumbnailUrlViewer(url) {
  if (!url) return "";

  const text = String(url).trim();

  if (text.indexOf("drive.google.com") !== -1) {
    const idMatch = text.match(/[-\w]{25,}/);

    if (idMatch) {
      return "https://drive.google.com/thumbnail?id=" + idMatch[0] + "&sz=w1000";
    }
  }

  return text;
}
