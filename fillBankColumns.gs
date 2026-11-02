// 総勘定元帳: 取引No.補完(AA・AB)と、N/A対応(AA)をまとめたスクリプト。
// 【AA・AB補完】同じ取引No.(A列)の行から、空欄の AA(銀行)・AB(銀行 該当ページ) を補完する。複数値は重複除去して改行連結。既存の値は上書きしない。

const KEY_COL = 1;            // A列 取引No
const TARGET_COLS = [27, 28]; // AA, AB

function onOpen() {
  SpreadsheetApp.getUi().createMenu('取引No.補完')
    .addItem('AA・ABを補完', 'fillBankColumns')
    .addItem('N/A対応', 'fillNA')
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

// ===== N/A対応: 特定の勘定科目の行は AA列 に「N/A」を入れる(AAに既に値がある行は上書きしない)。メニュー「取引No.補完 > N/A対応」から実行 =====
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
