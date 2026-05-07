/**
 * =========================================================================
 * 【重要】このビューアー（在庫閲覧専用サイドバー）を使うための切り替え手順
 * =========================================================================
 * *「サイドバー注文機能付き」から、この「サイドバー在庫閲覧用」に切り替えるためには、
 * サイドバーを開く関数（showInit）を1箇所だけ書き換える必要があります。
 * * 変更前： var html = HtmlService.createTemplateFromFile('Sidebar')
 *   変更後： var html = HtmlService.createTemplateFromFile('sidebar_view_only')
 * =========================================================================
 */

/**
 * サイドバー(ビューアー版)から呼び出される専用の関数
 * ★アキラさん究極設計【ハイブリッド版】★修正版
 * 基本情報（写真・名前等）はMASTERから直接取得し、在庫情報だけSKUからかき集める！
 */
function getSelectedImageUrlViewer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); 
  const range = sheet.getActiveCell();
  const row = range.getRow();
  const col = range.getColumn();
  
  const headerRow = 6; 
  if (row <= headerRow) return null; 

  // =========================================================
  // ① クリックした列が「04_」か確認し、MASTERシートの見出しを把握する
  // =========================================================
  const currentHeaders = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const clickedHeader = String(currentHeaders[col - 1] || "").trim();
  
  // クリックするのは「04_（写真表示）」の列で正解！
  if (!clickedHeader.startsWith("04_")) return null;

  const getMasterCol = (idPrefix) => {
    const idx = currentHeaders.findIndex(h => String(h).trim().startsWith(idPrefix));
    return idx !== -1 ? idx + 1 : null;
  };

  const colCode = getMasterCol("06_");
  if (!colCode) return null;

  // MASTERシートからクリックした行の商品コードを取得（先頭7文字）
  const rawCode = String(sheet.getRange(row, colCode).getValue() || "").trim();
  const itemCode = rawCode.substring(0, 7).toUpperCase(); 

  if (rawCode === "" || !itemCode) {
    return {
      url: "", urls: [], code: "-", nameEn: "-", nameJp: "商品コードがありません",
      priceVn: 0, priceJp: 0, siteUrl: "", stockData: []
    };
  }

  // =========================================================
  // ② 基本情報（写真・名前など）は「今クリックしているMASTERの行」から直接拾う！
  // =========================================================
  // ★大修正：URLの文字を引っこ抜くのは「04_」ではなく「05_」！！
  const colPhotoUrl = getMasterCol("05_"); 
  const colSiteUrl  = getMasterCol("03_");
  const colNameEn   = getMasterCol("08_") || getMasterCol("111_"); 
  const colNameJp   = getMasterCol("16_") || getMasterCol("121_");
  const colPriceVn  = getMasterCol("14_");
  const colPriceJp  = getMasterCol("15_");

  let primaryPhotoUrl = "";
  if (colPhotoUrl) {
    // 05_列のセルから、スマートチップまたは通常のURL文字を読み取る
    const richCell = sheet.getRange(row, colPhotoUrl).getRichTextValue();
    if (richCell) primaryPhotoUrl = richCell.getLinkUrl() || richCell.getText();
    if (!primaryPhotoUrl) primaryPhotoUrl = String(sheet.getRange(row, colPhotoUrl).getValue() || "");
  }

  const siteUrl = colSiteUrl ? String(sheet.getRange(row, colSiteUrl).getValue() || "") : "";
  const nameEn  = colNameEn ? String(sheet.getRange(row, colNameEn).getValue() || "") : "";
  const nameJp  = colNameJp ? String(sheet.getRange(row, colNameJp).getValue() || "") : "";
  const priceVn = colPriceVn ? Number(sheet.getRange(row, colPriceVn).getValue()) || 0 : 0;
  const priceJp = colPriceJp ? Number(sheet.getRange(row, colPriceJp).getValue()) || 0 : 0;

  // =========================================================
  // ③ SKUシートに行き、在庫・バリエーション情報「だけ」をかき集める
  // =========================================================
  const skuSheet = ss.getSheetByName("SKU");
  let stockData = [];
  
  if (skuSheet) {
    const skuHeaders = skuSheet.getRange(headerRow, 1, 1, skuSheet.getLastColumn()).getValues()[0];
    const getSkuCol = (idPrefix) => {
      const idx = skuHeaders.findIndex(h => String(h).trim().startsWith(idPrefix));
      return idx !== -1 ? idx + 1 : null;
    };

    const colSkuFullCode = getSkuCol("06_") || getSkuCol("061_"); 
    const colVar         = getSkuCol("11_");
    const colSize        = getSkuCol("09_");
    const colStock       = getSkuCol("52_"); // 在庫列

    if (colSkuFullCode) {
      const skuData = skuSheet.getDataRange().getValues();
      for (let i = headerRow; i < skuData.length; i++) {
        let skuFullCode = String(skuData[i][colSkuFullCode - 1] || "").trim();
        let baseCodeOfSku = skuFullCode.substring(0, 7).toUpperCase();
        
        // 先頭7文字が一致したら、サイズと在庫の情報をストックする
        if (baseCodeOfSku === itemCode && baseCodeOfSku !== "") {
          let currentStock = colStock ? Number(skuData[i][colStock - 1] || 0) : 0;
          stockData.push({
            variation: colVar ? String(skuData[i][colVar - 1] || "") : "基本",
            size: colSize ? String(skuData[i][colSize - 1] || "") : "-",
            stock: currentStock
          });
        }
      }
    }
  }

  // もしSKUシートにデータが1件もなかった場合のフォロー
  if (stockData.length === 0 && nameJp === "") {
    return {
      url: "", urls: [], code: itemCode, nameEn: "-", nameJp: "SKUシートに登録がありません",
      priceVn: 0, priceJp: 0, siteUrl: "", stockData: []
    };
  }

  // =========================================================
  // ④ 写真URLの変換（神ツール制限突破版）
  // =========================================================
  const allUrls = primaryPhotoUrl.split(/[,、\n\s]+/).map(u => u.trim()).filter(u => u !== "").map(getThumbnailUrlViewer);
  
  return {
    url: allUrls[0] || "",
    urls: allUrls, 
    code: itemCode,
    nameEn: nameEn,
    nameJp: nameJp, 
    priceVn: priceVn,
    priceJp: priceJp,
    siteUrl: siteUrl,
    stockData: stockData 
  };
}

/**
 * 画像変換用：アキラさん提供の「制限突破版（サムネイル方式）」
 */
function getThumbnailUrlViewer(url) {
  if (!url) return "";
  if (url.indexOf("drive.google.com") !== -1) {
    const idMatch = url.match(/[-\w]{25,}/);
    if (idMatch) {
      return "https://drive.google.com/thumbnail?id=" + idMatch[0] + "&sz=w1000";
    }
  }
  return url;
}