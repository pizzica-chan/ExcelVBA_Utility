Attribute VB_Name = "modRange"
Option Explicit

' セル範囲の取得、検索、値の転記、空白行の削除。
' このモジュールだけでインポートして使えます。
' 引数付きの Sub はマクロ一覧に出ません。Run_ で始まる Sub だけをボタンやショートカットに割り当てます。

' 【LastRow】最終行を返す。データが無いときは 0。
'   第2引数を省略すると、数式も含めたシート全体の最終行。
'   列は 2 でも "B" でも指定できる。
' 使用例:
'   finalRow = LastRow(ActiveSheet, "B")
'   If finalRow >= 2 Then
'       Set dataBody = ActiveSheet.Range("A2:C" & finalRow)
'   End If
' 解説: 今開いているシートの B 列で、最後に値がある行を探す。2 行目以降にデータがあれば、A2 から C 列のその行までを dataBody にする。見出しだけのときは範囲を作らない。
Public Function LastRow(Optional ByVal ws As Worksheet, Optional ByVal columnIndex As Variant) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    If IsMissing(columnIndex) Then
        LastRow = LastUsedRow(ws)
    Else
        LastRow = LastRowInColumn(ws, ColumnNumber(columnIndex))
    End If
End Function

' 【LastColumn】最終列を返す。データが無いときは 0。
'   行を省略するとシート全体。行を渡すとその行だけを見る。
' 使用例:
'   lastCol = LastColumn(ActiveSheet, 1)   ' 1行目の最終列
'   Debug.Print ColumnLetter(lastCol)      ' 例: "F"
' 解説: 1 行目だけを見て、右端に値がある列番号を lastCol に入れる。続けて列記号へ直し、イミディエイトに "F" のように出す。
Public Function LastColumn(Optional ByVal ws As Worksheet, Optional ByVal rowIndex As Variant) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    If IsMissing(rowIndex) Then
        LastColumn = LastUsedColumn(ws)
    Else
        LastColumn = LastColumnInRow(ws, CLng(rowIndex))
    End If
End Function

' 【DataRange】A1 から最終セルまでの外接範囲。データが無いときは Nothing。
'   途中の空白セルも範囲に含む。隣り合う表の塊だけなら CurrentDataRegion。
' 使用例:
'   Set body = DataRange(Worksheets("売上"))
'   If Not body Is Nothing Then body.Copy
' 解説: 「売上」シートの A1 から、データがある右下までの四角い範囲をコピーする。シートが空なら何もしない。
Public Function DataRange(Optional ByVal ws As Worksheet) As Range
    Dim r As Long
    Dim c As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    r = LastRow(ws)
    c = LastColumn(ws)
    If r = 0 Or c = 0 Then Exit Function
    Set DataRange = ws.Range(ws.Cells(1, 1), ws.Cells(r, c))
End Function

' 【CurrentDataRegion】指定セルを含む連続領域。Ctrl+Shift+* と同じ。
' 使用例:
'   Set table = CurrentDataRegion(ActiveCell)
'   table.AutoFilter
' 解説: 選択中のセルを含む、空白で区切られた一塊の表にフィルタを付ける。隣の別表までは含まない。
Public Function CurrentDataRegion(ByVal anchor As Range) As Range
    Set CurrentDataRegion = anchor.CurrentRegion
End Function

' 【ColumnLetter】列番号を列記号にする。1 → "A"、27 → "AA"。
' 使用例:
'   Debug.Print ColumnLetter(LastColumn())   ' 例: "H"
' 解説: シート全体の最終列を番号で取り、A や H のような列記号にしてイミディエイトへ出す。セルに =ColumnLetter(27) と書くと "AA" になる。
' ワークシート関数としても使える。=ColumnLetter(27)
Public Function ColumnLetter(ByVal columnIndex As Long) As String
    Dim n As Long
    Dim remainder As Long
    If columnIndex < 1 Or columnIndex > 16384 Then
        Err.Raise 5, "ColumnLetter", "列番号は 1 から 16384 までです。"
    End If
    n = columnIndex
    Do While n > 0
        remainder = (n - 1) Mod 26
        ColumnLetter = Chr$(65 + remainder) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

' 【ColumnNumber】列記号または列番号を、1 始まりの列番号にする。
' 使用例:
'   n = ColumnNumber("AA")   ' 27
'   n = ColumnNumber(3)      ' 3
' 解説: "AA" は 27 列目なので n は 27。数字の 3 を渡したときは、そのまま 3 列目として返す。セルでは =ColumnNumber("XFD") が最終列 16384 になる。
' ワークシート関数としても使える。=ColumnNumber("XFD")
Public Function ColumnNumber(ByVal columnRef As Variant) As Long
    Dim text As String
    Dim i As Long
    Dim ch As String
    If IsNumeric(columnRef) And Not IsEmpty(columnRef) Then
        ColumnNumber = CLng(columnRef)
        If ColumnNumber < 1 Or ColumnNumber > 16384 Then
            Err.Raise 5, "ColumnNumber", "列番号は 1 から 16384 までです。"
        End If
        Exit Function
    End If
    text = UCase$(Trim$(CStr(columnRef)))
    If Len(text) = 0 Then Err.Raise 5, "ColumnNumber", "列が指定されていません。"
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "A" Or ch > "Z" Then
            Err.Raise 5, "ColumnNumber", "列指定が不正です: " & CStr(columnRef)
        End If
        ColumnNumber = ColumnNumber * 26 + (Asc(ch) - 64)
    Next i
    If ColumnNumber < 1 Or ColumnNumber > 16384 Then
        Err.Raise 5, "ColumnNumber", "列指定が不正です: " & CStr(columnRef)
    End If
End Function

' 【CopyValues】値だけを転記する。書式と数式はコピーしない。
'   貼り付け先は左上セルを渡せば、元範囲と同じ大きさに書き込まれる。
' 使用例:
'   CopyValues Worksheets("元").Range("A1:C100"), Worksheets("先").Range("A1")
' 解説: 「元」の A1:C100 に見えている値を、「先」の A1 から同じ 100 行 3 列へ書き写す。数式や色はコピーしない。
Public Sub CopyValues(ByVal source As Range, ByVal destination As Range)
    If source.Cells.CountLarge = 1 Then
        destination.Cells(1, 1).Value = source.Value
    Else
        destination.Cells(1, 1).Resize(source.Rows.Count, source.Columns.Count).Value = source.Value
    End If
End Sub

' 【PasteValuesInPlace】範囲内の数式を、計算結果の値に置き換える。
' 使用例:
'   PasteValuesInPlace Worksheets("集計").Range("A2:F100")
' 解説: 「集計」の A2:F100 にある数式を、その場の計算結果に置き換える。あとから元データを消しても、この範囲の数字は残る。
Public Sub PasteValuesInPlace(ByVal target As Range)
    target.Value = target.Value
End Sub

' 【FindAll】一致したセルをすべて Collection で返す。既定では大文字小文字を区別しない。
'   lookAt は xlWhole（完全一致）または xlPart（部分一致）。
' 使用例:
'   Dim hit As Range
'   For Each hit In FindAll(Range("A:A"), "東京", xlPart)
'       Debug.Print hit.Address, hit.Value
'   Next hit
' 解説: A 列から「東京」を含むセルをすべて拾い、番地と中身をイミディエイトへ出す。完全に「東京」だけのセルに限るときは xlWhole を渡す。
Public Function FindAll(ByVal searchRange As Range, ByVal findWhat As Variant, _
    Optional ByVal lookAt As XlLookAt = xlWhole, _
    Optional ByVal matchCase As Boolean = False) As Collection

    Dim result As New Collection
    Dim found As Range
    Dim firstAddress As String
    Set found = searchRange.Find( _
        What:=findWhat, _
        After:=searchRange.Cells(searchRange.Rows.Count, searchRange.Columns.Count), _
        LookIn:=xlValues, _
        LookAt:=lookAt, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=matchCase)
    If Not found Is Nothing Then
        firstAddress = found.Address
        Do
            result.Add found
            Set found = searchRange.FindNext(found)
        Loop While Not found Is Nothing And found.Address <> firstAddress
    End If
    Set FindAll = result
End Function

' 【DeleteBlankRows】範囲内がすべて空欄の行を、シートの行ごと削除する。
'   同じ行の、範囲の外にある値も消える。連続した 50 万セル以内を指定する。
' 使用例:
'   DeleteBlankRows Worksheets("名簿").Range("A2:D500")
' 解説: 「名簿」の 2 行目から 500 行目のうち、A から D がすべて空の行をシートから消す。E 列以降に値があっても、その行ごと消える。
Public Sub DeleteBlankRows(ByVal target As Range)
    Dim data As Variant
    Dim rowIndex As Long
    Dim blockEnd As Long
    EnsureSingleReasonableRange target, "DeleteBlankRows"
    data = ToMatrix(target)
    blockEnd = 0
    For rowIndex = UBound(data, 1) To 1 Step -1
        If IsBlankDataRow(data, rowIndex) Then
            If blockEnd = 0 Then blockEnd = rowIndex
        ElseIf blockEnd <> 0 Then
            target.Rows(rowIndex + 1).Resize(blockEnd - rowIndex).EntireRow.Delete
            blockEnd = 0
        End If
    Next rowIndex
    If blockEnd <> 0 Then
        target.Rows(1).Resize(blockEnd).EntireRow.Delete
    End If
End Sub

' 【FillDownBlanks】空欄を、同じ列のすぐ上の値で埋める。
'   数式は値になる。先頭行の空欄はそのまま残る。
' 使用例:
'   ' A列の区分が下の空欄まで続いて見える表を、セルごとに埋める
'   FillDownBlanks Worksheets("伝票").Range("A2:A200")
' 解説: 「伝票」の A2:A200 で、空欄をすぐ上の区分で埋める。先頭の A2 が空なら、そのセルだけは空のまま残る。
Public Sub FillDownBlanks(ByVal target As Range)
    Dim data As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    EnsureSingleReasonableRange target, "FillDownBlanks"
    data = ToMatrix(target)
    For colIndex = 1 To UBound(data, 2)
        For rowIndex = 2 To UBound(data, 1)
            If IsBlankValue(data(rowIndex, colIndex)) Then
                data(rowIndex, colIndex) = data(rowIndex - 1, colIndex)
            End If
        Next rowIndex
    Next colIndex
    target.Value = data
End Sub

' 【UnmergeAndFill】結合を解除し、入っていた値を各セルへ埋める。
'   結合が範囲の外まではみ出していても、その結合全体を解除する。
' 使用例:
'   UnmergeAndFill Worksheets("帳票").Range("A1:H50")
' 解説: 「帳票」の A1:H50 に掛かる結合を外し、結合セルに入っていた文字を、分かれた各セルへ同じ値で入れる。
Public Sub UnmergeAndFill(ByVal target As Range)
    Dim seen As Object
    Dim cell As Range
    Dim area As Range
    Dim addressKey As Variant
    Dim value As Variant
    EnsureSingleReasonableRange target, "UnmergeAndFill"
    Set seen = CreateObject("Scripting.Dictionary")
    For Each cell In target.Cells
        If cell.MergeCells Then
            addressKey = cell.MergeArea.Address
            If Not seen.Exists(addressKey) Then
                seen.Add addressKey, cell.MergeArea.Cells(1, 1).Value
            End If
        End If
    Next cell
    For Each addressKey In seen.Keys
        Set area = target.Worksheet.Range(CStr(addressKey))
        value = seen(addressKey)
        area.UnMerge
        area.Value = value
    Next addressKey
End Sub

' 【HighlightDuplicates】重複する値に条件付き書式で色を付ける。空欄は対象外。
'   fillColor を省略すると黄色。RGB で別の色も指定できる。
' 使用例:
'   HighlightDuplicates Range("A2:A500")
'   HighlightDuplicates Range("C2:C500"), RGB(255, 199, 206)
' 解説: 1 行目は選択範囲の重複を黄色にする。2 行目は C2:C500 の重複を薄い赤にする。1 回しか出ない値と空欄には色が付かない。
Public Sub HighlightDuplicates(ByVal target As Range, Optional ByVal fillColor As Long = 65535)
    With target.FormatConditions.AddUniqueValues
        .DupeUnique = xlDuplicate
        .Interior.Color = fillColor
    End With
End Sub

' 【ShowAllData】オートフィルタの絞り込みを解除する。フィルタ矢印は残る。
' 使用例:
'   ShowAllData Worksheets("一覧")
' 解説: 「一覧」で隠れている行をすべて表示する。フィルタの矢印は残るので、あとから別の条件で絞り直せる。
Public Sub ShowAllData(ByVal ws As Worksheet)
    If ws Is Nothing Then Err.Raise 5, "ShowAllData", "シートが指定されていません。"
    If ws.FilterMode Then ws.ShowAllData
End Sub

' 【FilterByHeader】見出しセルの列で、表をその値に絞り込む。
'   見出しと同じ行から連続した範囲を表とみなす。
' 使用例:
'   ' 1行目が見出しで、C1 が「支社」のとき、東京だけ残す
'   FilterByHeader Worksheets("一覧").Range("C1"), "東京"
' 解説: C1 の見出しがある列を「東京」で絞る。見出し行から続く表以外は対象にしない。
Public Sub FilterByHeader(ByVal headerCell As Range, ByVal criteria As String)
    Dim dataRegion As Range
    Dim fieldIndex As Long
    Set dataRegion = headerCell.CurrentRegion
    fieldIndex = headerCell.Column - dataRegion.Column + 1
    dataRegion.AutoFilter Field:=fieldIndex, Criteria1:=criteria
End Sub

' 【Run_DeleteBlankRows】マクロ一覧用。選択範囲の空欄行を行ごと削除する。
'   実行前に、表の範囲を選択しておく。列全体の選択はエラーになる。
' 使用例:
'   ボタンにこのマクロを割り当て、A2:D500 を選択して実行する。
' 解説: 選んだ範囲がすべて空の行だけが消える。列全体を選ぶとエラーになるので、表の行範囲を選ぶ。
Public Sub Run_DeleteBlankRows()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    RunWithUi "DeleteBlankRows", target
End Sub

' 【Run_ConvertSelectionToValues】マクロ一覧用。選択範囲の数式を値にする。
' 使用例:
'   数式が入った集計範囲を選択してから実行する。
' 解説: 選んだセルの数式が、そのときの計算結果に置き換わる。元の数式には戻せないので、必要なブックは先に保存する。
Public Sub Run_ConvertSelectionToValues()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    RunWithUi "PasteValuesInPlace", target
End Sub

' 【Run_HighlightDuplicates】マクロ一覧用。選択範囲の重複値を黄色にする。
' 使用例:
'   社員番号の列を選択して実行すると、同じ番号のセルに色が付く。
' 解説: 2 回以上出ている値だけが黄色になる。空欄は重複として扱わない。
Public Sub Run_HighlightDuplicates()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    On Error GoTo EH
    HighlightDuplicates target
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "HighlightDuplicates"
End Sub

' 【Run_UnmergeAndFill】マクロ一覧用。選択範囲の結合を解除して値を埋める。
' 使用例:
'   結合セルのある帳票範囲を選択して実行する。
' 解説: 結合が外れ、表示されていた値が分かれたセルのそれぞれに入る。結合が選択の外まではみ出していても、その結合全体が外れる。
Public Sub Run_UnmergeAndFill()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    RunWithUi "UnmergeAndFill", target
End Sub

' 【Run_FillDownBlanks】マクロ一覧用。選択範囲の空欄を上の値で埋める。
' 使用例:
'   区分列で上の値が下の空欄まで続いている範囲を選択して実行する。
' 解説: 空欄が、同じ列のすぐ上の値で埋まる。範囲の先頭行が空のときは、そのセルは空のまま残る。
Public Sub Run_FillDownBlanks()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    RunWithUi "FillDownBlanks", target
End Sub

' 【Run_ClearFilter】マクロ一覧用。アクティブシートの絞り込みを解除する。
' 使用例:
'   フィルタで行が隠れているシートを表示してから実行する。
' 解説: 今開いているシートの絞り込みが外れ、隠れていた行がすべて見える。フィルタの矢印は残る。
Public Sub Run_ClearFilter()
    On Error GoTo EH
    ShowAllData ActiveSheet
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "ShowAllData"
End Sub

Private Sub RunWithUi(ByVal actionName As String, ByVal target As Range)
    Dim prevScreen As Boolean
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo EH
    Select Case actionName
        Case "DeleteBlankRows": DeleteBlankRows target
        Case "PasteValuesInPlace": PasteValuesInPlace target
        Case "UnmergeAndFill": UnmergeAndFill target
        Case "FillDownBlanks": FillDownBlanks target
    End Select
    Application.ScreenUpdating = prevScreen
    Exit Sub
EH:
    Application.ScreenUpdating = prevScreen
    MsgBox Err.Description, vbExclamation, actionName
End Sub

Private Function RequireSelection() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation
        Exit Function
    End If
    Set RequireSelection = Selection
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then LastUsedColumn = found.Column
End Function

Private Function LastRowInColumn(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    Dim rowIndex As Long
    rowIndex = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    If rowIndex = 1 Then
        If IsBlankCell(ws.Cells(1, columnIndex)) Then Exit Function
    End If
    LastRowInColumn = rowIndex
End Function

Private Function LastColumnInRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Long
    Dim columnIndex As Long
    If rowIndex < 1 Then Err.Raise 5, "LastColumn", "行番号が不正です。"
    columnIndex = ws.Cells(rowIndex, ws.Columns.Count).End(xlToLeft).Column
    If columnIndex = 1 Then
        If IsBlankCell(ws.Cells(rowIndex, 1)) Then Exit Function
    End If
    LastColumnInRow = columnIndex
End Function

Private Function IsBlankCell(ByVal cell As Range) As Boolean
    If cell.HasFormula Then Exit Function
    IsBlankCell = IsBlankValue(cell.Value)
End Function

Private Function IsBlankDataRow(ByVal data As Variant, ByVal rowIndex As Long) As Boolean
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If Not IsBlankValue(data(rowIndex, colIndex)) Then Exit Function
    Next colIndex
    IsBlankDataRow = True
End Function

Private Function IsBlankValue(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        IsBlankValue = IsEmpty(value) Or IsNull(value)
        Exit Function
    End If
    If VarType(value) = vbString Then
        IsBlankValue = (Len(Trim$(Replace(Replace(CStr(value), ChrW(&HA0), " "), ChrW(&H3000), " "))) = 0)
    End If
End Function

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

Private Sub EnsureSingleReasonableRange(ByVal target As Range, ByVal sourceName As String)
    If target Is Nothing Then Err.Raise 5, sourceName, "範囲が空です。"
    If target.Areas.Count <> 1 Then Err.Raise 5, sourceName, "連続した 1 つの範囲を指定してください。"
    If target.Rows.Count = target.Worksheet.Rows.Count Then
        Err.Raise 5, sourceName, "列全体ではなく、データの範囲を指定してください。"
    End If
    If target.Cells.CountLarge > 500000 Then
        Err.Raise 5, sourceName, "範囲が大きすぎます（50 万セルまで）。"
    End If
End Sub
