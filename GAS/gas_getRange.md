- **getRangeメソッドでセルやセル範囲を取得する**

- **getValueメソッドで単体セルの値を取得する**
- **getValuesメソッドでセル範囲の値を配列として取得する**

### **getValuesメソッドでセル範囲の値を取得する**

さて、**セル範囲の値をまとめて取得**したい場合には、**getValuesメソッド**を使う方法があります。getValue**s**ということで、複数形ですね。

書き方はこちらです。

Rangeオブジェクト.**getValues**()

例えば、以下のスクリプトを実行してみましょう。

    functionmyFunction() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getRange('A3:C4');
      console.log(range.getValues());
    }


## [シート全体の最後の行（最終行）を取得する方法](https://yokonoji.work/gas-get-last-row)

シート全体の最後の行（最終行）を取得するには、[getLastRowメソッド](https://developers.google.com/apps-script/reference/spreadsheet/sheet#getlastrow)を使用します。

```
var lastRow = sheet.getLastRow();
```

### 特定の列の最後の行（最終行）を取得する方法

各列の最終行を取得したい場合もあるかと思います。

そんな場合は、lastRow関数のような方法で指定した列の最終行を取得できます。

```js
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadsheet.getSheetByName("シート1");

function myFunction() {
  // シート全体の最終行を取得
  var lastRowAll = sheet.getLastRow();
  // 取得したい列を指定 A列=>1, B列=>2,・・・
  var column = 3;

  // 指定した列の最終行を取得
  var last_row = lastRow(lastRowAll, column);
  Logger.log(last_row);
}

function lastRow(lastRowAll, column) {
  var range = sheet.getRange(1, column, lastRowAll, column).getValues();
  for(var i=lastRowAll-1; i>=0; i--){
    if(range[i][0] != ""){
      return i + 1;
    }
  }
}
```

lastRow関数には「シート全体の最終行」と「最終行を取得したい列」を引数として渡しています。

これらの情報を元にして、下図のように空ではないセルを探します。そして、値があればその行が最終行となります。

---

## 最後の列（最終列）を取得する方法

スプレッドシートの最後の行（最終行）を取得するためには、対象のシートを指定しておく必要があります。

```js
// スプレッドシートを取得する
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
// シートを取得する
var sheet = spreadsheet.getSheetByName("シート1");
```

### シート全体の最後の列（最終列）を取得する方法

シート全体の最後の列（最終列）を取得するには、[getLastColumnメソッド](https://developers.google.com/apps-script/reference/spreadsheet/sheet#getlastcolumn)を使用します。

```js
var lastColumn = sheet.getLastColumn();
```

### 特定の行の最後の列（最終列）を取得する方法

各行の最終列を取得したい場合もあるかと思います。

```js
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadsheet.getSheetByName("シート1");

function myFunction() {
  // シート全体の最終列を取得
  var lastColumnAll = sheet.getLastColumn();
  // 取得したい行を指定 1行目=>1, 2行目=>2,・・・
  var row = 3;

  // 指定した列の最終行を取得
  var last_column = lastColumn(lastColumnAll, row);
  Logger.log(last_column);
}
function lastColumn(lastColumnAll, row) {
  var range = sheet.getRange(row, 1, row, lastColumnAll).getValues();
  for(var i=lastColumnAll-1; i>=0; i--){
    if(range[0][i] != ""){
      return i + 1;
    }
  }
}
```

lastColumn関数には「シート全体の最終列」と「最終列を取得したい行」を引数として渡しています。
