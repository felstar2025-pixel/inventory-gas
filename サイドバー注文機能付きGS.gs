// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット1】スプレッドシートのメニューを作る機能 ★進化★
// 「注文書を作成」の中に、国ごとのサブメニューを作りました！
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // ★国ごとの注文書ボタンをまとめた「サブメニュー」を作る
  var poMenu = ui.createMenu('📄 注文書を作成')
    .addItem('🇯🇵 日本円 (¥)', 'createPO_JPY')
    .addItem('🇻🇳 ベトナムドン (VND)', 'createPO_VND')
    .addItem('🇨🇳 中国元 (元) ※準備中', 'createPO_CNY'); // 将来用

  // メインのメニューを組み立てる
  ui.createMenu('★プレビュー')
    .addItem('サイドバー表示', 'showInit')
    .addSubMenu(poMenu) // ★ここにサブメニューを合体！
    .addItem('🛒 カートを空にする', 'resetCart')
    .addToUi();
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット2】カートを空っぽ（リセット）にする機能
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function resetCart() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('確認', '本当にカートを空（リセット）にしますか？\n（F〜M列の7行目以降のデータが消去されます）', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("CART");
    if (sheet) {
      var lastRow = Math.max(sheet.getLastRow(), 7); 
      sheet.getRange(7, 6, lastRow - 6 + 1, 8).clearContent(); 
      SpreadsheetApp.flush();
      ui.alert('カートをリセットしました！');
    }
  }
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット3】右側にサイドバーを出す機能
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function showInit() {
  var html = HtmlService.createTemplateFromFile('sidebar_view_only')
    .evaluate()
    .setTitle('商品情報')
    .setWidth(350); 
  SpreadsheetApp.getUi().showSidebar(html);
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット4】カートの合計数・金額を計算する機能たち
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function getInitialTotals() {
  try {
    // サイドバーを開いた瞬間にマスターと同期して、カートの値段を最新にする
    return refreshAndCalculateTotals();
  } catch(e) { return null; }
}

function calculateTotals(sheet) {
  if (!sheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) sheet = ss.getSheetByName("CART");
  }
  if (!sheet) return { count: 0, vnd: 0, jpy: 0, discount: 0 };

  SpreadsheetApp.flush(); 
  var jpyTotal = sheet.getRange("O2").getValue() || 0; 
  var vndTotal = sheet.getRange("N2").getValue() || 0; 
  var discount = sheet.getRange("N3").getValue() || 0; 
  var totalQty = sheet.getRange("L2").getValue() || 0; 

  if (discount > 0 && discount < 1) {
    discount = Math.round(discount * 100);
  }
  return { count: totalQty, vnd: vndTotal, jpy: jpyTotal, discount: discount };
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット5】サイズを正しい順番に並べ替えるためのルール辞書
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
var sizeOrder = {"XS":1, "S":2, "M":3, "L":4, "XL":5, "フリー":6, "F":6, "FREE":6};
function getOrder(val) {
  if(!val) return 99; 
  var clean = val.toString().toUpperCase().trim();
  return sizeOrder[clean] || 99; 
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット6】選択した商品の「画像や詳細データ」を探す機能
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function getSelectedImageUrl() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var currentSheet = ss.getActiveSheet();
    var cell = currentSheet.getActiveCell();
    
    if (cell.getRow() < 7) return null;

// 1. コードを読み取る（F列から取得）
    var rawCode = cell.offset(0, 2).getValue().toString().trim();

// 2. もしコードに「-（ハイフン）」が入っていたら、左側だけを取り出す（マスター検索用）
// これで "FF10001AA-A" が "FF10001AA" になり、マスターと一致するようになります！
    var code = rawCode.split('-')[0];
    if (!code) return null;

    var markCell = "";
    if (cell.getColumn() > 1) {
      markCell = cell.offset(0, -1).getValue().toString();
    }
    var isSoldOutAll = markCell.indexOf('完') !== -1;

    var masterSheet = ss.getSheetByName("MASTER");
    var masterData = masterSheet.getDataRange().getValues();
    
    var res = null;
    for (var i = 1; i < masterData.length; i++) {
      if (masterData[i][5] == code) { 
        var rawLines = masterData[i][4] ? masterData[i][4].toString().split('\n') : [];
        var processedUrls = [];

        rawLines.forEach(function(line) {
          var u = line.trim();
          if (!u) return; 
          var finalUrl = u;
          if (u.indexOf("drive.google.com") !== -1) {
            var match = u.match(/[-\w]{25,}/);
            if (match) finalUrl = "https://drive.google.com/thumbnail?id=" + match[0] + "&sz=w500";
          }
          processedUrls.push(finalUrl);
        });

        res = {
          url: processedUrls[0] || "", 
          urls: processedUrls,         
          code: code,
          soldOutAll: isSoldOutAll, 
          nameEn:  masterData[i][6],  
          nameJp:  masterData[i][7],  
          size:    masterData[i][8],  
          colorVn: masterData[i][9],  
          colorJp: masterData[i][10], 
          cCodes:  masterData[i][11], 
          priceVn: masterData[i][12], 
          priceJp: masterData[i][13],

                    // === 【別ウィンドウで問屋サイトを開く処理：データ準備】 ===
          // ・MASTERシートの O列（15列目＝プログラム上は[14]）からURLを読み込んでいます
          // ・もしURLをP列にするなら[15]、Q列にするなら[16]に変更してください
          siteUrl: masterData[i][2] ? masterData[i][2].toString() : ""
        };  

        break; 
      }
    }
    if (res == null) return null;

    var cartSheet = ss.getSheetByName("CART");
    var inCartData = []; 
    if (cartSheet) {
      var cartData = cartSheet.getDataRange().getValues();
      for (var c = 6; c < cartData.length; c++) {
        if (cartData[c][5] == code) { 
          inCartData.push({
            color: cartData[c][9], 
            size:  cartData[c][10], 
            qty:   cartData[c][11]  
          });
        }
      }
    }
    inCartData.sort(function(a, b) {
      if (a.color !== b.color) return a.color < b.color ? -1 : 1;
      var oA = getOrder(a.size);
      var oB = getOrder(b.size);
      return oA - oB;
    });
    res.inCartData = inCartData;
    return res;
  } catch (e) { return null; }
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット7】カートに商品を追加する機能（超高速版）
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function addToCart(selectedInfo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cartSheet = ss.getSheetByName("CART") || ss.insertSheet("CART");
    
    var START_ROW = 7; 
    var fValues = cartSheet.getRange("F1:F").getValues();
    var lastRowInF = 0;
    for (var i = fValues.length - 1; i >= 0; i--) {
      if (fValues[i][0] !== "") { lastRowInF = i + 1; break; }
    }
    
    var addQty = parseInt(selectedInfo.quantity) || 1; 
    if (addQty <= 0) return { totals: calculateTotals(cartSheet) }; 
    
    var foundRow = -1;

    if (lastRowInF >= START_ROW) {
      var cartData = cartSheet.getRange(1, 1, lastRowInF, 17).getValues();
      for (var j = START_ROW - 1; j < cartData.length; j++) {
        if (cartData[j][5] == selectedInfo.code && 
            cartData[j][10] == selectedInfo.selectedSize && 
            cartData[j][9] == selectedInfo.selectedColorJp) {
          foundRow = j + 1;
          break;
        }
      }
    }

    if (foundRow > 0) {
      var qtyCell = cartSheet.getRange(foundRow, 12);
      var currentQty = parseInt(qtyCell.getValue()) || 0;
      qtyCell.setValue(currentQty + addQty);
    } else {
      var nextRow = Math.max(lastRowInF + 1, START_ROW);
      var vns = selectedInfo.colorVn.split(',');
      var jps = selectedInfo.colorJp.split(',');
      var idx = jps.findIndex(function(c){ return c.trim() === selectedInfo.selectedColorJp; });
      var finalColorVn = (idx !== -1 && vns[idx]) ? vns[idx].trim() : "";

      var newRowData = [[
        selectedInfo.code, 
        selectedInfo.nameEn, 
        selectedInfo.nameJp, 
        finalColorVn, 
        selectedInfo.selectedColorJp, 
        selectedInfo.selectedSize, 
        addQty, 
        selectedInfo.priceVn
      ]];
      cartSheet.getRange(nextRow, 6, 1, 8).setValues(newRowData);
      customSortCart(cartSheet, START_ROW, nextRow); 
    }
    // 追加した直後にも最新価格をチェックする
    return { totals: refreshAndCalculateTotals() };
  } catch (e) { return { message: "エラー: " + e.toString() }; }
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット8】カートから商品を減らす・消す機能
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function removeFromCart(selectedInfo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cartSheet = ss.getSheetByName("CART");
    if (!cartSheet) return { error: "カートがありません" };

    var START_ROW = 7;
    var fValues = cartSheet.getRange("F1:F").getValues();
    var lastRowInF = 0;
    for (var i = fValues.length - 1; i >= 0; i--) {
      if (fValues[i][0] !== "") { lastRowInF = i + 1; break; }
    }
    if (lastRowInF < START_ROW) return { totals: calculateTotals(cartSheet) };

    var removeQty = parseInt(selectedInfo.quantity) || 1;
    var foundRow = -1;
    var cartData = cartSheet.getRange(1, 1, lastRowInF, 17).getValues();

    for (var j = START_ROW - 1; j < cartData.length; j++) {
      if (cartData[j][5] == selectedInfo.code && 
          cartData[j][10] == selectedInfo.selectedSize && 
          cartData[j][9] == selectedInfo.selectedColorJp) {
        foundRow = j + 1;
        break;
      }
    }

    if (foundRow > 0) {
      var qtyCell = cartSheet.getRange(foundRow, 12);
      var currentQty = parseInt(qtyCell.getValue()) || 0;
      var newQty = currentQty - removeQty;

      if (newQty <= 0) {
        cartSheet.deleteRow(foundRow);
      } else {
        qtyCell.setValue(newQty);
      }
    }
    // 減らした直後にも最新価格をチェックする
    return { totals: refreshAndCalculateTotals() };
  } catch (e) { return { message: "エラー: " + e.toString() }; }
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット9】カートの並べ替え＆★最新価格シンクロ機能★
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function refreshAndCalculateTotals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CART");
  var masterSheet = ss.getSheetByName("MASTER"); // ★追加：マスターシートも呼ぶ
  if (!sheet || !masterSheet) return null;

  var START_ROW = 7;
  var lastRow = sheet.getLastRow();
  
  if (lastRow >= START_ROW) {
    // --- 【ステップ1】マスターシートの「最新価格表」を裏側で作る ---
    var masterData = masterSheet.getDataRange().getValues();
    var priceDict = {};
    for (var m = 1; m < masterData.length; m++) {
      var mCode = masterData[m][5]; // マスターのF列(コード)
      if (mCode) {
        // コードを鍵にして、最新の値段(M列)を記憶しておく
        priceDict[mCode.toString().trim()] = masterData[m][12]; 
      }
    }

    // --- 【ステップ2】カートのデータを読み込む ---
    var range = sheet.getRange(START_ROW, 6, lastRow - START_ROW + 1, 8); 
    var data = range.getValues();
    var newData = [];

    // --- 【ステップ3】カートの中身を1つずつチェック ---
    for (var i = 0; i < data.length; i++) {
      var qty = parseInt(data[i][6]) || 0; 
      var code = data[i][0] ? data[i][0].toString().trim() : "";
      
      if (qty > 0 && code !== "") {
        // ★ ここが魔法！マスターに同じコードがあれば、値段だけ最新に書き換える！
        if (priceDict[code]) {
          data[i][7] = priceDict[code]; // M列（単価VND）を最新価格で上書き
        }
        newData.push(data[i]); // 整理用ボックスに入れる
      }
    }

    // --- 【ステップ4】綺麗になった最新データをシートに戻す ---
    range.clearContent();
    if (newData.length > 0) {
      sheet.getRange(START_ROW, 6, newData.length, 8).setValues(newData);
      customSortCart(sheet, START_ROW, START_ROW + newData.length - 1);
    }
  }
  return calculateTotals(sheet);
}

function customSortCart(sheet, startRow, lastRow) {
  if (lastRow < startRow) return;
  var range = sheet.getRange(startRow, 6, lastRow - startRow + 1, 8); 
  var data = range.getValues();

  data.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1; 
    if (a[4] !== b[4]) return a[4] < b[4] ? -1 : 1; 
    var oA = getOrder(a[5]); 
    var oB = getOrder(b[5]);
    if (oA !== oB) return oA - oB;                  
    return a[5] < b[5] ? -1 : 1; 
  });
  range.setValues(data);
}

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット10-A】各国のボタンが押された時のスイッチ ★NEW★
// メニューで選ばれた国に合わせて、通貨マークの指示を出します。
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function createPO_JPY() { runCreatePO("JPY"); }
function createPO_VND() { runCreatePO("VND"); }
function createPO_CNY() { runCreatePO("CNY"); } // 将来用

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
// 【ユニット10-B】PO（注文書）を作成する本体機能 ★自動切替・完璧版★
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function runCreatePO(currencyType) {
  
  // 1. 通貨によって「マーク」と「合計金額の場所」を切り替える設定
  var markFormat = '#,##0'; 
  var totalCell = "O2";     // デフォルト(円)の合計
  var discountCell = "O4";  // デフォルト(円)の割引額
  var rateCell = "O3";      // デフォルト(円)の割引率

  if (currencyType === "JPY") {
    markFormat = '¥#,##0';
    totalCell = "O2";
    discountCell = "O4";
    rateCell = "O3";
  } else if (currencyType === "VND") {
    markFormat = '#,##0" ₫"';
    totalCell = "N2";
    discountCell = "N4";
    rateCell = "N3";
  } else if (currencyType === "CNY") {
    markFormat = '#,##0" 元"';
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cartSheet = ss.getSheetByName("CART");
  var poSheet = ss.getSheetByName("PO");
  
  // 2. POシートの古い明細（18行目以降）を真っ白にお掃除
  var lastRowPO = poSheet.getLastRow();
  if (lastRowPO >= 18) {
    poSheet.getRange(18, 2, lastRowPO - 17, 8).clearContent(); 
    poSheet.getRange(18, 2, lastRowPO - 17, 8).setBackground(null); 
  }
  
  // 3. 固定データをPOシートへ運ぶ
  poSheet.getRange("C5").setValue(cartSheet.getRange("H4").getValue());  // お客さんの名前
  poSheet.getRange("F16").setValue(cartSheet.getRange(totalCell).getValue()); // ★合計（自動切替）
  poSheet.getRange("E16").setValue(cartSheet.getRange(rateCell).getValue());  // ★割引率（自動切替）
  poSheet.getRange("G16").setValue(cartSheet.getRange(discountCell).getValue()); // ★割引額（自動切替）
  
  poSheet.getRange("H9").setValue(cartSheet.getRange("L2").getValue());  // ★追加：合計購入数！
  
  // ★追加＆修正：H7（グランドトータル）と I16（最終合計）にお化粧をする
  poSheet.getRange("H7").setNumberFormat(markFormat);
  poSheet.getRange("I16").setNumberFormat(markFormat);
  
  // 4. CARTの明細データの読み込み (7行目スタート)
  var CART_START_ROW = 7; 
  var cartLastRow = cartSheet.getLastRow();
  if (cartLastRow < CART_START_ROW) {
    Browser.msgBox("カートに商品がありません！"); 
    return;
  }
  
  var cartData = cartSheet.getRange(CART_START_ROW, 1, cartLastRow - CART_START_ROW + 1, 16).getValues(); 
  var poData = [];      
  var colorRules = [];  
  var currentColor = "#ffffff"; 
  var previousCode = "";        
  var rowCount = 1; 
  
  for (var i = 0; i < cartData.length; i++) {
    var row = cartData[i];
    var code = row[5]; // F列: 商品コード
    if (!code || code === "") continue; 
    
    if (previousCode !== "" && code !== previousCode) {
      currentColor = (currentColor === "#ffffff") ? "#f3f3f3" : "#ffffff";
    }
    
    // ★通貨によって「どの列のデータを運ぶか」を自動で切り替える！
    var itemName, itemColor, itemSize, itemQty, itemPrice, itemSubtotal;
    
    itemSize = row[10]; // K列 (共通)
    itemQty = row[11];  // L列 (共通)

    if (currencyType === "VND") {
      itemName = row[6];      // G列: 商品名(ベトナム語)
      itemColor = row[8];     // I列: カラー(ベトナム語)
      itemPrice = row[12];    // M列: 単価(VND)
      itemSubtotal = row[13]; // N列: 小計(VND)
    } else {
      itemName = row[7];      // H列: 商品名(日本語)
      itemColor = row[9];     // J列: カラー(日本語)
      itemPrice = row[14];    // O列: 単価(JPY)
      itemSubtotal = row[15]; // P列: 小計(JPY)
    }
    
    poData.push([
      rowCount, code, itemName, itemColor, itemSize, itemQty, itemPrice, itemSubtotal
    ]);
    
    colorRules.push([currentColor, currentColor, currentColor, currentColor, currentColor, currentColor, currentColor, currentColor]);
    previousCode = code; 
    rowCount++;
  }
  
  // 5. 明細データを貼り付け＆色塗り
  if (poData.length > 0) {
    var targetRange = poSheet.getRange(18, 2, poData.length, 8); 
    targetRange.setValues(poData);        
    targetRange.setBackgrounds(colorRules); 
    
    poSheet.getRange(18, 2, poData.length, 1).setNumberFormat("0");      
    poSheet.getRange(18, 8, poData.length, 2).setNumberFormat('#,##0');  
  }
  
  // 6. 自動で「POシート」の画面に切り替える！
  poSheet.activate();
}


/**
 * [神ツール・制限突破版] 
 * Googleドライブの厳しい制限をすり抜ける裏ワザURL変換を搭載したハイブリッド版です！
 */
function embedImagesFromUrlList() {
  // --- [設定] ---
  var urlSheetName = "URLリスト"; 
  var urlRange = "A1:A33";       
  var destSheetName = "MASTER";   
  var destStartCell = "D7";       
  // --------------

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var urlSheet = ss.getSheetByName(urlSheetName);
  var destSheet = ss.getSheetByName(destSheetName);

  if (!urlSheet || !destSheet) {
    SpreadsheetApp.getUi().alert("シート名が間違っているかもしれません。");
    return;
  }

  var urls = urlSheet.getRange(urlRange).getValues();
  var destRange = destSheet.getRange(destStartCell);
  var destStartRow = destRange.getRow();
  var destCol = destRange.getColumn();

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput('<div style="padding: 20px;">最新の裏ワザで画像をセット中です...<br>そのままお待ちください。</div>'),
    "処理中..."
  );

  for (var i = 0; i < urls.length; i++) {
    var url = urls[i][0].toString().trim().replace(/^"|"$/g, '');

    if (url && url.match(/^http/)) {
      try {
        if (url.indexOf("drive.google.com") !== -1) {
          // ★ここが突破口！ 
          // 古い「uc?export」ではなく、ブロックされにくい「サムネイル生成用URL」に変換します
          var idMatch = url.match(/[-\w]{25,}/);
          if (idMatch) {
            var driveUrl = "https://drive.google.com/thumbnail?id=" + idMatch[0] + "&sz=w1000";
            
            var imageValue = SpreadsheetApp.newImageValue()
                                            .withSourceUrl(driveUrl)
                                            .withAltText("商品画像")
                                            .build();
            destSheet.getRange(destStartRow + i, destCol).setValue(imageValue);
          }
        } else {
          // 外部サイトの場合はIMAGE関数でサクッと表示
          var formula = '=IMAGE("' + url + '")';
          destSheet.getRange(destStartRow + i, destCol).setFormula(formula);
        }
      } catch (e) {
        console.error("エラー: " + e);
        destSheet.getRange(destStartRow + i, destCol).setValue("画像エラー");
      }
    }
  }

  SpreadsheetApp.getUi().alert("全件の画像セットが完了しました！(制限突破版)");
}