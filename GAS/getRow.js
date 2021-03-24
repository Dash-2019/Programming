function myFunction() {
//アクティブなスプレッドシートのシートを取得する
let mySheet = SpreadsheetApp.getActiveSheet();
//選択されているアクティブなセルを取得する
let myActiveCell = mySheet.getActiveCell();
//アクティブなセルからRow:行、Column:列を取得する
let selectedRow = myActiveCell.getRow();
let selectedColumn = myActiveCell.getColumn();
//スプレッドシート上でアクティブなセルをポップアップ表示
Browser.msgBox("セルの選択位置","行："+selectedRow+ "、列："+selectedColumn, Browser.Buttons.OK);
}
