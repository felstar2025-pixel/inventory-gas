/**
 * 【トランアリー社専用】
 * 商品ごとに S と M のセット数を確認し、割引額を計算する魔法の関数
 */
function calculateTranAliDiscount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cartSheet = ss.getSheetByName("CART"); // ★シート名を「CART」に固定しました
  
  if (!cartSheet) return "シートが見つかりません";

  // --- 【設定：今のCARTシートの列に合わせました】 ---
  // getRange(7, 6, ...) で「6列目（F列）」から読み込むので、そこを 0番目 として数えます
  const COL_CODE  = 0; // F列 (商品コード)
  const COL_SIZE  = 5; // K列 (サイズ)
  const COL_QTY   = 6; // L列 (個数)

  const lastRow = cartSheet.getLastRow();
  if (lastRow < 7) return 0; // 7行目より下が空なら 0円
  
  // 7行目、6列目(F列)から、11列分(P列まで)のデータを一気に取得
  const data = cartSheet.getRange(7, 6, lastRow - 6, 11).getValues();
  let itemSummary = {}; 

  // --- 【1. 商品（品名コード）ごとに集計】 ---
  data.forEach(row => {
    let code  = row[COL_CODE];
    let size  = String(row[COL_SIZE]).toUpperCase().trim();
    let count = parseInt(row[COL_QTY]) || 0; 

    if (code) {
      // まだ集計用の箱がない商品コードなら新しく作る
      if (!itemSummary[code]) {
        itemSummary[code] = { S: 0, M: 0, totalPcs: 0 };
      }
      // Sの数、Mの数、そしてその商品全体の数をカウント
      if (size === 'S') itemSummary[code].S += count;
      if (size === 'M') itemSummary[code].M += count;
      itemSummary[code].totalPcs += count;
    }
  });

  let totalDiscountVND = 0;

  // --- 【2. 割引ランク判定ロジック（トランアリー流）】 ---
  for (let code in itemSummary) {
    let s = itemSummary[code].S;
    let m = itemSummary[code].M;
    let totalPcs = itemSummary[code].totalPcs;

    // SとM、どちらか少ない方の数が「セット数」になる
    let setPairs = Math.min(s, m);
    
    let unitDiscount = 90000; // 基本（バラ）の値引き単価

    // 揃っているセット数に応じて、その商品全体の値引き単価をランクアップ！
    if (setPairs >= 10) {
      unitDiscount = 130000; // 10S 10M 揃った
    } else if (setPairs >= 4) {
      unitDiscount = 110000; // 4S 4M 揃った
    } else if (setPairs >= 2) {
      unitDiscount = 100000; // 2S 2M 揃った
    }

    // 「決定した値引き単価」 × 「その商品の全着数」 を加算
    totalDiscountVND += (totalPcs * unitDiscount);
  }

  // --- 【3. 結果を返す】 ---
  // あきらさんの希望通り、マイナス（-）をつけて返します
  return -totalDiscountVND;
}