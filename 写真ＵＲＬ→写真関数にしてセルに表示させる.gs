/**
 * 現在選択している行の「写真URL(05_)」を「写真表示(04_)」に即座に変換する
 * ※ボタンに登録して使うと最高に便利です
 */
function refreshSelectedImage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRow = sheet.getActiveCell().getRow();
  
  if (activeRow < 7) {
    SpreadsheetApp.getUi().alert("7行目以降の行を選択してください。");
    return;
  }

  // 今のシートのヘッダーを読み取って列を特定
  const headers = sheet.getRange(6, 1, 1, sheet.getLastColumn()).getValues()[0];
  let col04 = -1, col05 = -1;
  headers.forEach((h, i) => {
    const cleanH = String(h).replace(/＿/g, "_").replace(/\s/g, "");
    if (cleanH.startsWith("04_")) col04 = i + 1;
    if (cleanH.startsWith("05_")) col05 = i + 1;
  });

  if (col04 === -1 || col05 === -1) {
    SpreadsheetApp.getUi().alert("04_写真 または 05_写真URL の列が見つかりません。");
    return;
  }

  const url = sheet.getRange(activeRow, col05).getValue();
  if (!url) {
    SpreadsheetApp.getUi().alert("URLが入っていません。");
    return;
  }

  // 変換して貼り付け
  const directUrl = getDirectImageUrl(url);
  sheet.getRange(activeRow, col04).setFormula(`=IMAGE("${directUrl}")`);
  
  // 完了を通知（右下に小さく出す）
  ss.toast("写真を更新しました！", "画像変換完了");
}