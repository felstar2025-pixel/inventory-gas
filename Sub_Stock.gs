function generateStockMatrix() {
  // === 【設定エリア】あきらさんと決めた完璧なルール ===
  var masterSheetName = "MASTER"; 
  var stockSheetName = "在庫管理";  
  
  var startRow = 7;       // データ開始行
  var colCode   = 6 - 1;  // F列（商品コード）
  var colNameVN = 7 - 1;  // G列（ベトナム品名）
  var colNameJP = 8 - 1;  // H列（日本品名）
  var colSize   = 9 - 1;  // I列（サイズ）
  var colColor  = 11 - 1; // K列（日本の色）
  
  // 固定で並べる8種類のサイズ（9行ブロック用）
  var fixedSizes = ["S", "M", "L", "XL", "XXL", "XS", "フリー", "予備"];
  // ==========================================

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(masterSheetName);
  
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("エラー：「" + masterSheetName + "」が見つかりません！");
    return;
  }

  // 「在庫管理」シートがなければ作り、あれば一旦まっさらにする（何度でもやり直し可能！）
  var stockSheet = ss.getSheetByName(stockSheetName);
  if (!stockSheet) {
    stockSheet = ss.insertSheet(stockSheetName);
  } else {
    stockSheet.clear(); 
  }

  var lastRow = masterSheet.getLastRow();
  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert("エラー：7行目以降にデータがありません！");
    return;
  }
  
  var data = masterSheet.getRange(startRow, 1, lastRow - startRow + 1, masterSheet.getLastColumn()).getValues();
  
  var finalData = []; 
  var maxColumns = 4; // A〜D列までは固定なので最低4列

  for (var i = 0; i < data.length; i++) {
    var productCode = data[i][colCode];
    if (!productCode || productCode.toString().trim() === "") continue;

    // 品名の取得
    var nameVN = data[i][colNameVN] ? data[i][colNameVN].toString().trim() : "";
    var nameJP = data[i][colNameJP] ? data[i][colNameJP].toString().trim() : "";

    // 色とサイズをカンマで切り分け
    var colorString = data[i][colColor] ? data[i][colColor].toString() : "色なし";
    var colors = colorString.split(",").map(function(c) { return c.trim(); });

    var sizeString = data[i][colSize] ? data[i][colSize].toString() : "";
    var productSizes = sizeString.split(",").map(function(s) { return s.trim(); });

    if (colors.length + 4 > maxColumns) {
      maxColumns = colors.length + 4;
    }

    // --- 🔨 9行ブロックの組み立て開始 ---
    
    // 1行目： [商品コード, 日本の品名, ベトナムの品名, 空白, 色1, 色2...]
    var headerRow = [productCode, nameJP, nameVN, ""];
    headerRow = headerRow.concat(colors);
    finalData.push(headerRow);

    // 2〜9行目： サイズの展開
    for (var s = 0; s < fixedSizes.length; s++) {
      var currentSize = fixedSizes[s];
      
      // A(コード), B(日本品名), C(ベトナム品名)は空白、D列にサイズ名
      var sizeRow = ["", "", "", currentSize]; 
      
      // マスターのI列にこのサイズが存在するか？
      var isValidSize = (productSizes.indexOf(currentSize) !== -1);

      for (var c = 0; c < colors.length; c++) {
        if (isValidSize) {
          sizeRow.push("");  // 存在するサイズなら「空白（後で0を入力用）」
        } else {
          sizeRow.push("-"); // 存在しないサイズなら「-」でブロック！
        }
      }
      finalData.push(sizeRow);
    }
  }

  // --- 📦 横幅調整 ---
  for (var r = 0; r < finalData.length; r++) {
    while (finalData[r].length < maxColumns) {
      finalData[r].push(""); 
    }
  }

  // データを一気に貼り付け
  if (finalData.length > 0) {
    stockSheet.getRange(1, 1, finalData.length, maxColumns).setValues(finalData);
    SpreadsheetApp.getUi().alert("大成功！品名入り・サイズ判定付きの在庫マトリックスが完成しました！");
  } else {
    SpreadsheetApp.getUi().alert("データが見つかりませんでした。");
  }
}