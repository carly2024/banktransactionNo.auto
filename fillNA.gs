// 総勘定元帳: 特定の勘定科目の行は AA列 に「N/A」を入れる。AAに既に値がある行は上書きしない。
// メニュー「取引No.補完 > N/A対応」から実行(メニューは fillBankColumns.gs の onOpen で作成)。

const NA_TARGET_COL = 27;          // AA列
const NA_ACCOUNT_HEADER = '勘定科目'; // 1行目の見出し名(この列で勘定科目を判定)
const NA_ACCOUNTS = ['為替差益', '為替差損', 'EV売上', 'デジタルサイネージ売上'];

function fillNA() {
  const sh = getLedgerSheet();
  if (!sh) {
    SpreadsheetApp.getUi().alert('「総勘定元帳」を含む名前の表示中のシートが見つかりません。');
    return;
  }
  const last = sh.getLastRow();
  if (last < 2) return;
  const n = last - 1;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const accIdx = headers.indexOf(NA_ACCOUNT_HEADER);
  if (accIdx < 0) {
    SpreadsheetApp.getUi().alert('1行目に「' + NA_ACCOUNT_HEADER + '」の見出しが見つかりません。');
    return;
  }
  const accounts = sh.getRange(2, accIdx + 1, n, 1).getValues().map(r => String(r[0]).trim());
  const rng = sh.getRange(2, NA_TARGET_COL, n, 1);
  const vals = rng.getValues();
  const forms = rng.getFormulas();

  let filled = 0;
  accounts.forEach((a, r) => {
    if (NA_ACCOUNTS.indexOf(a) >= 0 && forms[r][0] === '' && String(vals[r][0]).trim() === '') {
      vals[r][0] = 'N/A';
      filled++;
    }
  });

  // 数式セルは数式のまま書き戻す
  rng.setValues(vals.map((row, r) => [forms[r][0] || row[0]]));
  SpreadsheetApp.getActive().toast(filled + ' セルに N/A を入れました', 'N/A対応', 5);
}
