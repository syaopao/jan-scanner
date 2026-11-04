/**
 * JANスキャナー連携 GAS
 * シート2 C:H 参照、シート1 A:D へ送信
 *
 * [設定] > [スクリプト プロパティ]
 * SPREADSHEET_ID = 対象スプレッドシートID
 * SCANNER_KEY = 自分で作る長いランダムな合言葉
 * 
 * デプロイ: ウェブアプリ、実行ユーザー=自分、アクセス=全員
 * 公開リポジトリに SCANNER_KEY を記載しないこと。
 */
function scannerConfig_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  var key = props.getProperty('SCANNER_KEY');
  if (!id || !key || key.length < 16) throw new Error('SPREADSHEET_ID / SCANNER_KEY（16文字以上）を設定してください');
  return { ss: SpreadsheetApp.openById(id), key: key };
}
function scannerAuth_(params) {
  var conf = scannerConfig_();
  if (!params || String(params.key || '') !== conf.key) throw new Error('認証キーが一致しません');
  return conf.ss;
}
function scannerJsonp_(callback, payload) {
  if (!/^janCb_[A-Za-z0-9_]{1,80}$/.test(String(callback || ''))) {
    return ContentService.createTextOutput('invalid callback').setMimeType(ContentService.MimeType.TEXT);
  }
  var body = JSON.stringify(payload).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return ContentService.createTextOutput(callback + '(' + body + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function scannerValidCode_(code) {
  return /^[A-Za-z0-9][A-Za-z0-9 ._\/-]{0,63}$/.test(String(code || ''));
}
function doGet(e) {
  var p = (e && e.parameter) || {};
  var callback = p.callback;
  try {
    var ss = scannerAuth_(p);
    var action = String(p.action || '');
    if (action === 'ping') return scannerJsonp_(callback, {ok:true, message:'接続成功'});
    if (action === 'status') {
      var requestId = String(p.requestId || '');
      if (!/^[A-Za-z0-9_-]{8,80}$/.test(requestId)) throw new Error('requestIdが不正です');
      var saved = CacheService.getScriptCache().get('receipt_' + requestId);
      return scannerJsonp_(callback, {ok:true, complete:!!saved, receipt:saved?JSON.parse(saved):null});
    }
    if (action !== 'lookup') throw new Error('未対応の処理です');
    var codes = String(p.codes || '').split(',').map(function(c){return c.trim();}).filter(scannerValidCode_);
    codes = [...new Set(codes)].slice(0,60);
    var results = {};
    if (!codes.length) return scannerJsonp_(callback, {ok:true, products:results});
    var sheet = ss.getSheetByName('シート2');
    if (!sheet) throw new Error('シート2がありません');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return scannerJsonp_(callback, {ok:true, products:results});
    var rows = sheet.getRange(2,3,lastRow-1,6).getDisplayValues();
    var wanted = {};
    codes.forEach(function(code){ wanted[code]=true; });
    for (var i=0;i<rows.length;i++) {
      var row = rows[i], sheetCode=String(row[0]||'').trim();
      if (!sheetCode) continue;
      // UPC-A と先頭0付きEAN-13を同一商品として検索するための照合
      var requestCodes = [sheetCode];
      if (sheetCode.length===13 && sheetCode[0]==='0') requestCodes.push(sheetCode.slice(1));
      if (sheetCode.length===12) requestCodes.push('0'+sheetCode);
      var match = requestCodes.filter(function(code){return wanted[code] && !results[code];});
      if (!match.length) continue;
      var image = '';
      try {
        var imageCell = sheet.getRange(i+2,5);
        var value = imageCell.getValue();
        if (value && value.valueType === SpreadsheetApp.ValueType.IMAGE) image = value.getContentUrl();
        else {
          var formula = imageCell.getFormula();
          var found = formula && formula.match(/^=IMAGE\(\s*"([^"]+)"/i);
          if (found) image = found[1];
          else if (typeof value === 'string' && /^https:\/\//.test(value)) image=value;
        }
      } catch (ignoreImage) {}
      var product = {
        code:sheetCode,
        title:String(row[1]||''),
        category:String(row[3]||''),
        market:String(row[4]||''),
        sold:String(row[5]||''),
        image:/^https:\/\//i.test(image)?image:''
      };
      match.forEach(function(code){results[code]=product;});
      if (Object.keys(results).length>=codes.length) break;
    }
    return scannerJsonp_(callback,{ok:true,products:results});
  } catch(err) {
    return scannerJsonp_(callback,{ok:false,error:String(err && err.message || err)});
  }
}
function doPost(e) {
  var response;
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var ss = scannerAuth_(data);
    if (data.action !== 'save') throw new Error('保存処理ではありません');
    var requestId = String(data.requestId || '');
    if (!/^[A-Za-z0-9_-]{8,80}$/.test(requestId)) throw new Error('requestIdが不正です');
    var entries = data.items;
    if (!Array.isArray(entries) || !entries.length || entries.length > 200) throw new Error('送信件数が範囲外です');
    var rows = entries.map(function(item) {
      var code = String(item.code || '');
      var qty = Number(item.qty);
      if (!scannerValidCode_(code) || !Number.isInteger(qty) || qty < 1 || qty > 100000) throw new Error('バーコードまたは数量が不正です');
      return [new Date(), false, code, qty];
    });
    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      var cache = CacheService.getScriptCache(), cacheKey='receipt_'+requestId;
      var prior = cache.get(cacheKey);
      if (prior) response = JSON.parse(prior);
      else {
        var sheet = ss.getSheetByName('シート1');
        if (!sheet) throw new Error('シート1がありません');
        var firstRow = Math.max(2,sheet.getLastRow()+1);
        // バーコードの先頭0が数値変換で落ちないよう書き込み前に書式設定
        sheet.getRange(firstRow,3,rows.length,1).setNumberFormat('@');
        sheet.getRange(firstRow,1,rows.length,4).setValues(rows);
        sheet.getRange(firstRow,1,rows.length,1).setNumberFormat('yyyy/MM/dd HH:mm:ss');
        SpreadsheetApp.flush();
        response={saved:rows.length,firstRow:firstRow};
        cache.put(cacheKey,JSON.stringify(response),21600);
      }
    } finally {lock.releaseLock();}
    return ContentService.createTextOutput(JSON.stringify({ok:true,receipt:response})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err && err.message || err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
