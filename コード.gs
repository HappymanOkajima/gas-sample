/**
 * アプリの説明
 *
 * 1. 概要
 * スプレッドシートの「テンプレート」シートに、メールの件名と本文を登録します。
 * 件名と本文には変数を利用できます。 ${} に囲まれた値が、customersシートの内容で置換されます（後述）。
 *
 * スプレッドシートの「customers」シートに、メールの送信先アドレスや変数の値を登録します。
 * 件名や本文中の変数を置換するために使われます。
 
 * また、送信可能フラグを設定できます。
 * 送信可能フラグが 0 の場合、メールを送信しません。
 *
 *
 * 2. 操作方法
 *   1. スプレッドシートの「テンプレート」シートに、メールの件名と本文を登録します。
 *   2. スプレッドシートの「customers」シートに、メールの送信先アドレスや変数の値を登録します。
 *   3. スクリプトエディタで run 関数を選択し実行します。
 */

// スプレッドシートのIDは適宜置き換えてください。
var spreadSheet = SpreadsheetApp.openById('xxxxxx');
var template    = spreadSheet.getSheetByName('テンプレート');
var customers   = spreadSheet.getSheetByName('customers').getDataRange().getValues();

function run() {
  var subject = template.getRange('B1').getValue();
  var body    = template.getRange('B2').getValue();
  
  for(var i = 1; i < customers.length; i++) {
    var customer = customers[i];
    
    if(customer[4] === 0) return;
    
    var vars = [
      {
        before: 'UserName',
        after : customer[0]
      },
      {
        before: 'StaffName',
        after : customer[2]
      },
      {
        before: 'StaffTel',
        after : customer[3]
      }
    ];
    
    var _subject = replaceTemplate_(subject, vars);
    var _body = replaceTemplate_(body, vars);
    
    var emailAddress = customer[1];
    GmailApp.sendEmail(emailAddress, _subject, _body, {
      noReply: true
    });
  }
}

function replaceTemplate_(template, vars) {
  var _template = template;
  
  for(var i = 0; i < vars.length; i++) {
    var _var = vars[i];
    var regExp = new RegExp('\\${' + _var.before + '}', 'g');
    _template = _template.replace(regExp, _var.after);
  }
  
  return _template;
}