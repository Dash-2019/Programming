'計算

Sub Arraysample()
    Application.ScreenUpdating = False
    Dim i As Long
    Dim fen As Double
    Dim Table As Variant

    Table = Range("A2:C99")

    For i = 1 To 98
      If Table(i, 1) <> Empty And Table(i, 1) <> 0 And Table(i, 1) <> "--" Then
                 
            fen = Table(i, 1) - Table(i, 2)
            Table(i, 3) = Round(fen / Table(i, 1), 5)
         End If      
    Next

    Range("A2:C99") = Table

End Sub


'*-----------------------------------------------------------------------
Sub Vlookup_Dic()
    Dim start_time As Double
    Dim 列位置 As Integer
    Dim Sh1 As Worksheet, Sh2 As Worksheet
    Set Sh1 = Worksheets("Google")
    Set Sh2 = Worksheets("Daily-V")
    
    Dim Rng検索値 As Range
    Dim Rng検索範囲 As Range
    Dim Rng出力範囲 As Range
    
    Set Rng検索値 = Sh2.Range("B3:B110")
    Set Rng検索範囲 = Sh1.Range("B2:F99")
    Set Rng出力範囲 = Sh2.Range("G3:G110")
        列位置 = 5
  
    Application.ScreenUpdating = False
  
    Call Sample_Dic(Rng検索値, Rng検索範囲, 列位置, Rng出力範囲)
    Application.ScreenUpdating = True
End Sub

'...Dictionaryを使う
Sub Sample_Dic(ByVal Rng検索値 As Range, _
            ByVal Rng検索範囲 As Range, _
            ByVal 列位置 As Integer, _
            ByVal Rng出力範囲 As Range)
    Dim i As Long
    Dim ary()
    Dim myDic As New Dictionary
    For i = 1 To Rng検索範囲.Rows.Count
        If Not myDic.Exists(Rng検索範囲(i, 1).Value) Then
            myDic.Add Rng検索範囲(i, 1).Value, Rng検索範囲(i, 1).Offset(, 列位置 - 1).Value
        End If
    Next
    ReDim Preserve ary(1 To Rng出力範囲.Rows.Count, 1 To 2)
    For i = 1 To Rng検索値.Rows.Count
        ary(i, 1) = myDic.Item(Rng検索値(i, 1).Value)
    Next
    Rng出力範囲.Value = ary
End Sub


Sub Vlookup_Dic()
    Dim start_time As Double
    Dim 列位置 As Integer
    Dim Sh1 As Worksheet, Sh2 As Worksheet
    Set Sh1 = Worksheets("Google")
    Set Sh2 = Worksheets("Daily-V")
    
    Dim Rng検索値 As Range
    Dim Rng検索範囲 As Range
    Dim Rng出力範囲 As Range
    
    Set Rng検索値 = Sh2.Range("B3:B110")
    Set Rng検索範囲 = Sh1.Range("B2:F99")
    Set Rng出力範囲 = Sh2.Range("G3:G110")
        列位置 = 5
  
    Application.ScreenUpdating = False
  
    Call Sample_Dic(Rng検索値, Rng検索範囲, 列位置, Rng出力範囲)
    Application.ScreenUpdating = True
End Sub
'*-----------------------------------------------------------------------
'...Dictionaryを使う
Sub Sample_Dic(ByVal Rng検索値 As Range, _
            ByVal Rng検索範囲 As Range, _
            ByVal 列位置 As Integer, _
            ByVal Rng出力範囲 As Range)
    Dim i As Long
    Dim ary()
    Dim myDic As New Dictionary
    For i = 1 To Rng検索範囲.Rows.Count
        If Not myDic.Exists(Rng検索範囲(i, 1).Value) Then
            myDic.Add Rng検索範囲(i, 1).Value, Rng検索範囲(i, 1).Offset(, 列位置 - 1).Value
        End If
    Next
    ReDim Preserve ary(1 To Rng出力範囲.Rows.Count, 1 To 2)
    For i = 1 To Rng検索値.Rows.Count
        ary(i, 1) = myDic.Item(Rng検索値(i, 1).Value)
    Next
    Rng出力範囲.Value = ary
End Sub
