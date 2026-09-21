// Google Apps Script：接收報名表單，依身分寫入「數位部」或「其他單位」工作表
// 部署方式：
//   1. 新建一個 Google 試算表 > 擴充功能 > Apps Script > 貼上本檔全部內容
//   2. 部署 > 新增部署作業 > 類型「網頁應用程式」
//      執行身分「我」、誰可以存取「所有人」> 部署
//   3. 複製「網頁應用程式網址」，貼到 index.html 的 GAS_URL

const HEADERS = ['報名時間', '姓名', 'Email', '科別／單位', '自備筆電', 'Google 帳號', '期待／問題'];

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#b7e1cd');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const group = d.group === '數位部' ? '數位部' : '其他單位';
    const dept = group === '數位部' ? d.dept : d.unit;
    if (!d.name || !d.email || !dept) throw new Error('缺少必填欄位');
    getSheet_(group).appendRow([
      new Date(), d.name, d.email, dept, d.laptop || '', d.gaccount || '', d.expect || ''
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// 瀏覽器直接開網址時回報名人數，方便確認部署成功
function doGet() {
  const count = function (name) { return Math.max(getSheet_(name).getLastRow() - 1, 0); };
  return json_({ ok: true, 數位部: count('數位部'), 其他單位: count('其他單位') });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
