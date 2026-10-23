// 総勘定元帳: 同じ取引No.(A列)の行から、空欄の AA(銀行)・AB(銀行 該当ページ) を補完する。
// 同じ取引No.で値が複数ある場合は、重複を除いて改行で連結する。既存の値は上書きしない。

const KEY_COL = 1;            // A列 取引No
const TARGET_COLS = [27, 28]; // AA, AB

function onOpen() {
  SpreadsheetApp.getUi().createMenu('取引No.補完')
    .addItem('AA・ABを補完', 'fillBankColumns')
    .addToUi();
}

// 非表示でない、名前に「総勘定元帳」を含むシートを対象にする(例: 「総勘定元帳」「★総勘定元帳」)
function getLedgerSheet() {
  return SpreadsheetApp.getActive().getSheets()
    .find(s => !s.isSheetHidden() && s.getName().includes('総勘定元帳'));
}

function fillBankColumns() {
  const sh = getLedgerSheet();
  if (!sh) {
    SpreadsheetApp.getUi().alert('「総勘定元帳」を含む名前の表示中のシートが見つかりません。');
    return;
  }
  const last = sh.getLastRow();
  if (last < 2) return;
  const n = last - 1;
  const keys = sh.getRange(2, KEY_COL, n, 1).getValues().map(r => String(r[0]).trim());
  const rng = sh.getRange(2, TARGET_COLS[0], n, 2);
  const vals = rng.getValues();
  const forms = rng.getFormulas();
  const isEmpty = (r, c) => forms[r][c] === '' && String(vals[r][c]).trim() === '';

  let filled = 0;
  [0, 1].forEach(c => {
    const groups = {};
    keys.forEach((k, r) => {
      if (!k) return;
      const v = String(vals[r][c]).trim();
      if (forms[r][c] === '' && v) {
        const g = groups[k] = groups[k] || [];
        if (g.indexOf(v) < 0) g.push(v);
      }
    });
    keys.forEach((k, r) => {
      if (k && groups[k] && isEmpty(r, c)) {
        vals[r][c] = groups[k].join('\n');
        filled++;
      }
    });
  });

  // 数式セルは数式のまま書き戻す
  const out = vals.map((row, r) => row.map((v, c) => forms[r][c] || v));
  rng.setValues(out);
  SpreadsheetApp.getActive().toast(filled + ' セルを補完しました', '取引No.補完', 5);
}
