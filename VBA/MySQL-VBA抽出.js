Sub 抽出connectmySQL()
  Dim adoCon As Object ' ADOコネクション
  Dim adoRs As Object  ' ADOレコードセット
  Dim SQL As String    ' SQL
  Dim i As Long
  ' ADOコネクションを作成
  Set adoCon = CreateObject("ADODB.Connection")
  On Error GoTo ErrorTrap

  ' ODBC接続
  ' DRIVER={MySQL ODBC 8.0 Unicode Driver}
  adoCon.Open _
      "DRIVER={MySQL ODBC 5.1 Driver};" & _
      " SERVER=localhost;" & _
      " DATABASE=test;" & _
      " UID=root;" & _
      " PWD=1907;"

  MsgBox "DB接続成功"
' SQL文
  SQL = "SELECT * FROM fruit_table"

  ' SQLの実行
  Set adoRs = adoCon.Execute(SQL)

  ' レコードセット内の全ての行の読込が
  ' 終了するまで処理を繰り返す
  i = 2
  Do Until adoRs.EOF
    Cells(i, 1) = adoRs!ID
    Cells(i, 2) = adoRs!Name


    i = i + 1
    ' 次の行に移動する
    adoRs.MoveNext
  Loop

  ' 解放処理
  adoRs.Close
  adoCon.Close
  Set adoRs = Nothing
  Set adoCon = Nothing
Exit Sub

ErrorTrap:
    Set adoRs = Nothing
    Set adoCon = Nothing
    MsgBox (Err.Description)

End Sub
