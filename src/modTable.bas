Attribute VB_Name = "modTable"
Option Explicit

' Excel テーブル（リストオブジェクト）の作成、追記、読み取り。
' テーブル名はブック内で一意です。このモジュールだけでインポートして使えます。

' 【HasTable】ブック内にその名前のテーブルがあるか。テーブル名はブックで一意。
' 使用例:
'   If HasTable("売上テーブル") Then ClearTableRows FindTable("売上テーブル")
' 解説: ブックに「売上テーブル」があるときだけ、そのデータ行を消して見出しだけ残す。無いときに消そうとしてエラーにならない。
Public Function HasTable(ByVal tableName As String, Optional ByVal wb As Workbook) As Boolean
    HasTable = Not FindTable(tableName, wb) Is Nothing
End Function

' 【FindTable】テーブルを返す。無いときは Nothing。
' 使用例:
'   Set table = FindTable("売上テーブル")
'   If Not table Is Nothing Then AppendTableRow table, Array(Date, "東京", 100)
' 解説: 「売上テーブル」があれば、今日の日付・東京・100 の 1 行を末尾に足す。テーブルが無いときは何もしない。
Public Function FindTable(ByVal tableName As String, Optional ByVal wb As Workbook) As ListObject
    Dim ws As Worksheet
    Dim table As ListObject
    If wb Is Nothing Then Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        For Each table In ws.ListObjects
            If StrComp(table.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = table
                Exit Function
            End If
        Next table
    Next ws
End Function

' 【EnsureTable】範囲をテーブルにして返す。すでにテーブルならそれを返す。
'   先頭行を見出しにする。tableName を渡すとその名前にする。
' 使用例:
'   Set table = EnsureTable(Worksheets("売上").Range("A1:D20"), "売上テーブル")
' 解説: 「売上」の A1:D20 を、見出し付きの「売上テーブル」にする。すでにテーブルなら、新しく作らずそのテーブルを返す。
Public Function EnsureTable(ByVal source As Range, Optional ByVal tableName As String = "") As ListObject
    Dim existing As ListObject
    On Error Resume Next
    Set existing = source.Cells(1, 1).ListObject
    On Error GoTo 0
    If Not existing Is Nothing Then
        If Len(tableName) > 0 Then existing.Name = tableName
        Set EnsureTable = existing
        Exit Function
    End If
    Set EnsureTable = source.Worksheet.ListObjects.Add(xlSrcRange, source, , xlYes)
    On Error Resume Next
    EnsureTable.TableStyle = "TableStyleMedium2"
    On Error GoTo 0
    If Len(tableName) > 0 Then EnsureTable.Name = tableName
End Function

' 【TableColumnNumber】見出し名から列番号を返す。無いときは 0。大文字小文字は区別しない。
' 使用例:
'   qtyCol = TableColumnNumber(table, "数量")
' 解説: 見出しが「数量」の列が左から何番目かを qtyCol に入れる。無いときは 0。大文字小文字は区別しない。
Public Function TableColumnNumber(ByVal table As ListObject, ByVal columnName As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(CStr(column.Name), columnName, vbTextCompare) = 0 Then
            TableColumnNumber = column.Index
            Exit Function
        End If
    Next column
End Function

' 【ClearTableRows】テーブルのデータを消し、見出しだけ残す。合計行も外す。
' 使用例:
'   ClearTableRows FindTable("売上テーブル")
' 解説: 「売上テーブル」のデータ行をすべて消し、見出し行だけを残す。合計行が出ていれば先に外す。
Public Sub ClearTableRows(ByVal table As ListObject)
    Dim rowIndex As Long
    If table.ShowTotals Then table.ShowTotals = False
    For rowIndex = table.ListRows.Count To 1 Step -1
        table.ListRows(rowIndex).Delete
    Next rowIndex
End Sub

' 【AppendTableRow】テーブルの末尾に 1 行足す。
'   values は Array("東京", 3, 1200) のような 1 次元配列か、先頭列へ入れる単一値。
'   列数を超えた分は無視する。
' 使用例:
'   AppendTableRow FindTable("売上テーブル"), Array(Date, "大阪", 2, 800)
' 解説: 「売上テーブル」の末尾に、今日・大阪・2・800 の 1 行を足す。列の並びはテーブルの左からの順に合わせる。
Public Sub AppendTableRow(ByVal table As ListObject, ByVal values As Variant)
    Dim newRow As ListRow
    Dim lower As Long
    Dim upper As Long
    Dim offset As Long
    Dim colIndex As Long
    Set newRow = table.ListRows.Add
    If Not IsArray(values) Then
        newRow.Range.Cells(1, 1).Value = values
        Exit Sub
    End If
    lower = LBound(values)
    upper = UBound(values)
    For offset = lower To upper
        colIndex = offset - lower + 1
        If colIndex > table.ListColumns.Count Then Exit For
        newRow.Range.Cells(1, colIndex).Value = values(offset)
    Next offset
End Sub

' 【ReadTable】テーブルを二次元配列で返す。合計行は含めない。
'   includeHeader が False のときはデータ行だけ。データが無ければ Empty。
' 使用例:
'   data = ReadTable(FindTable("売上テーブル"), False)
'   If IsArray(data) Then Debug.Print data(1, 1)
' 解説: 「売上テーブル」のデータ行だけを配列で読み、先頭行の左端の値をイミディエイトへ出す。データが無ければ何も出さない。
Public Function ReadTable(ByVal table As ListObject, Optional ByVal includeHeader As Boolean = True) As Variant
    Dim body As Variant
    Dim header As Variant
    Dim combined() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim bodyRows As Long
    If table.DataBodyRange Is Nothing Then
        If includeHeader Then
            ReadTable = ToMatrix(table.HeaderRowRange)
        Else
            ReadTable = Empty
        End If
        Exit Function
    End If
    body = ToMatrix(table.DataBodyRange)
    If Not includeHeader Then
        ReadTable = body
        Exit Function
    End If
    header = ToMatrix(table.HeaderRowRange)
    bodyRows = UBound(body, 1)
    ReDim combined(1 To bodyRows + 1, 1 To table.ListColumns.Count)
    For colIndex = 1 To table.ListColumns.Count
        combined(1, colIndex) = header(1, colIndex)
        For rowIndex = 1 To bodyRows
            combined(rowIndex + 1, colIndex) = body(rowIndex, colIndex)
        Next rowIndex
    Next colIndex
    ReadTable = combined
End Function

' 【FilterTable】見出し名の列でテーブルを絞り込む。
' 使用例:
'   FilterTable FindTable("売上テーブル"), "支社", "東京"
' 解説: 「売上テーブル」の「支社」列が東京の行だけを表示する。見出し名が無いときはエラーになる。
Public Sub FilterTable(ByVal table As ListObject, ByVal columnName As String, ByVal criteria As String)
    Dim fieldIndex As Long
    fieldIndex = TableColumnNumber(table, columnName)
    If fieldIndex = 0 Then Err.Raise 5, "FilterTable", "列が見つかりません: " & columnName
    table.Range.AutoFilter Field:=fieldIndex, Criteria1:=criteria
End Sub

' 【ClearTableFilter】テーブルの絞り込みを解除する。フィルタボタンは残る。
' 使用例:
'   ClearTableFilter FindTable("売上テーブル")
' 解説: 「売上テーブル」の絞り込みを外し、隠れていた行をすべて表示する。フィルタボタンは残る。
Public Sub ClearTableFilter(ByVal table As ListObject)
    On Error Resume Next
    If table.ShowAutoFilter Then table.AutoFilter.ShowAllData
    On Error GoTo 0
End Sub

Private Function ToMatrix(ByVal source As Range) As Variant
    Dim data() As Variant
    If source.Cells.CountLarge = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = source.Value
        ToMatrix = data
    Else
        ToMatrix = source.Value
    End If
End Function
