Attribute VB_Name = "modClean"
Option Explicit

' 文字列のトリム、数値テキストの変換、重複行の削除、ハイパーリンクとメモの解除。
' このモジュールだけでインポートして使えます。

' 【TrimCells】文字列セルの前後の空白、NBSP、全角スペースを整える。
'   数式は変えない。途中の連続スペースはそのまま残す。
' 使用例:
'   TrimCells Worksheets("名簿").Range("A2:D500")
' 解説: 「名簿」の A2:D500 で、文字列の前後にある半角空白・全角スペース・NBSP を除く。数式セルと、文字の途中の連続スペースは変えない。
Public Sub TrimCells(ByVal target As Range)
    Dim textCells As Range
    Dim area As Range
    Dim cell As Range
    Dim cleaned As String
    Set textCells = TextConstants(target)
    If textCells Is Nothing Then Exit Sub
    For Each area In textCells.Areas
        For Each cell In area.Cells
            cleaned = NormalizeCellText(CStr(cell.Value))
            If cleaned <> CStr(cell.Value) Then cell.Value = cleaned
        Next cell
    Next area
End Sub

' 【FixNumbersStoredAsText】"123" のような数値テキストを数値にする。
'   "0123" のように先頭が 0 のコードは文字列のまま残す。数式は変えない。
' 使用例:
'   FixNumbersStoredAsText Worksheets("取込").Range("C2:C500")
' 解説: 「取込」の C2:C500 で、"123" のような数値に見える文字列を本当の数値にする。"0123" のように先頭が 0 のコードは文字列のまま残す。
Public Sub FixNumbersStoredAsText(ByVal target As Range)
    Dim textCells As Range
    Dim area As Range
    Dim cell As Range
    Dim text As String
    Set textCells = TextConstants(target)
    If textCells Is Nothing Then Exit Sub
    For Each area In textCells.Areas
        For Each cell In area.Cells
            text = Trim$(CStr(cell.Value))
            If IsConvertibleNumber(text) Then cell.Value = Val(text)
        Next cell
    Next area
End Sub

' 【ReplaceText】範囲内の文字列を置換する。既定では大文字小文字を区別しない。
' 使用例:
'   ReplaceText Worksheets("名簿").Range("A:A"), "（株）", "株式会社"
' 解説: 「名簿」の A 列で、「（株）」を「株式会社」に置き換える。大文字小文字は区別しない。セルの一部だけ一致しても置き換わる。
Public Sub ReplaceText(ByVal target As Range, ByVal findWhat As String, ByVal replaceWith As String, _
    Optional ByVal matchCase As Boolean = False)

    target.Replace What:=findWhat, Replacement:=replaceWith, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=matchCase
End Sub

' 【DeduplicateRows】重複行を削除し、削除した行数を返す。先頭行は見出しとして残す。
'   削除されるのは行全体なので、範囲の外にある同じ行の値も消える。
' 使用例:
'   removed = DeduplicateRows(Worksheets("名簿").Range("A1:D500"), xlYes)
'   MsgBox removed & " 行を削除しました。"
' 解説: 「名簿」の A1:D500 で、見出しを残して同じ内容の行を消す。消した行数をメッセージで出す。消えるのは行全体なので、E 列以降の同じ行も消える。
Public Function DeduplicateRows(ByVal target As Range, Optional ByVal header As XlYesNoGuess = xlYes) As Long
    Dim before As Long
    Dim cols() As Variant
    Dim i As Long
    If target Is Nothing Then Err.Raise 5, "DeduplicateRows", "範囲が空です。"
    before = FilledRowCount(target, header)
    If target.Columns.Count = 1 Then
        target.RemoveDuplicates Columns:=Array(1), Header:=header
    Else
        ReDim cols(1 To target.Columns.Count)
        For i = 1 To target.Columns.Count
            cols(i) = i
        Next i
        target.RemoveDuplicates Columns:=(cols), Header:=header
    End If
    DeduplicateRows = before - FilledRowCount(target, header)
End Function

Private Function FilledRowCount(ByVal target As Range, ByVal header As XlYesNoGuess) As Long
    Dim values As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim startRow As Long
    Dim rowFilled As Boolean
    startRow = 1
    If header = xlYes And target.Rows.Count > 1 Then startRow = 2
    If target.Cells.CountLarge = 1 Then
        If startRow = 1 And Not IsBlankCell(target.Value) Then FilledRowCount = 1
        Exit Function
    End If
    values = target.Value
    For rowIndex = startRow To UBound(values, 1)
        rowFilled = False
        For colIndex = 1 To UBound(values, 2)
            If Not IsBlankCell(values(rowIndex, colIndex)) Then
                rowFilled = True
                Exit For
            End If
        Next colIndex
        If rowFilled Then FilledRowCount = FilledRowCount + 1
    Next rowIndex
End Function

Private Function IsBlankCell(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        IsBlankCell = True
    ElseIf VarType(value) = vbString Then
        IsBlankCell = (Len(value) = 0)
    End If
End Function

' 【RemoveHyperlinks】範囲内のハイパーリンクを外す。表示文字は残る。
' 使用例:
'   RemoveHyperlinks Worksheets("名簿").Range("A2:A500")
' 解説: 「名簿」の A2:A500 からリンクだけを外す。セルに見えている文字はそのまま残る。
Public Sub RemoveHyperlinks(ByVal target As Range)
    target.Hyperlinks.Delete
End Sub

' 【ClearNotes】従来のメモを消す。対応している Excel では新しいコメントも消す。
' 使用例:
'   ClearNotes Worksheets("名簿").UsedRange
' 解説: 「名簿」で使っている範囲のメモを消す。新しい形式のコメントにも対応している Excel では、それも消える。
Public Sub ClearNotes(ByVal target As Range)
    target.ClearComments
    On Error Resume Next
    CallByName target, "ClearCommentsThreaded", VbMethod
    On Error GoTo 0
End Sub

' 【ClearConditionalFormats】範囲に付いている条件付き書式をすべて消す。
' 使用例:
'   ClearConditionalFormats Worksheets("一覧").Range("A2:A500")
' 解説: 「一覧」の A2:A500 に付いている条件付き書式をすべて消す。セルの値と、普通の塗りつぶしはそのまま。
Public Sub ClearConditionalFormats(ByVal target As Range)
    target.FormatConditions.Delete
End Sub

' 【Run_TrimSelection】マクロ一覧用。選択範囲の文字列の前後空白を除く。
' 使用例:
'   名簿の範囲を選択して実行する。数式セルは変わらない。
' 解説: 選んだ範囲の文字列から、前後の空白だけが除かれる。数式と、文字の途中のスペースは残る。
Public Sub Run_TrimSelection()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    On Error GoTo EH
    TrimCells target
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "TrimCells"
End Sub

' 【Run_FixNumericText】マクロ一覧用。選択範囲の数値テキストを数値にする。
'   0123 のような先頭ゼロのコードは文字列のまま残る。
' 使用例:
'   CSV 取り込みで数値が文字列になった列を選択して実行する。
' 解説: 選んだ範囲の "123" が数値の 123 になる。0123 のように先頭が 0 のコードは、コードのまま残る。
Public Sub Run_FixNumericText()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    On Error GoTo EH
    FixNumbersStoredAsText target
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "FixNumbersStoredAsText"
End Sub

' 【Run_RemoveDuplicates】マクロ一覧用。選択範囲の重複行を削除する。
'   先頭行は見出しとして残す。終わると削除件数を表示する。
' 使用例:
'   見出しを含む表全体を選択して実行する。
' 解説: 先頭行を見出しとして残し、同じ内容の行を削除する。終わると、削除した行数がメッセージで出る。
Public Sub Run_RemoveDuplicates()
    Dim target As Range
    Dim removed As Long
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    On Error GoTo EH
    removed = DeduplicateRows(target, xlYes)
    MsgBox CStr(removed) & " 行の重複を削除しました。先頭行は見出しとして残しています。", vbInformation, "DeduplicateRows"
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "DeduplicateRows"
End Sub

' 【Run_RemoveHyperlinks】マクロ一覧用。選択範囲のハイパーリンクを外す。
' 使用例:
'   リンクが付いたセル範囲を選択して実行する。文字はそのまま残る。
' 解説: 選んだ範囲のハイパーリンクだけが外れる。表示されている文字は変わらない。
Public Sub Run_RemoveHyperlinks()
    Dim target As Range
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    On Error GoTo EH
    RemoveHyperlinks target
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "RemoveHyperlinks"
End Sub

Private Function RequireSelection() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation
        Exit Function
    End If
    Set RequireSelection = Selection
End Function

Private Function TextConstants(ByVal target As Range) As Range
    On Error Resume Next
    Set TextConstants = target.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
End Function

Private Function NormalizeCellText(ByVal text As String) As String
    Dim ch As String
    text = Replace(text, ChrW(&HA0), " ")
    text = Replace(text, ChrW(&H3000), " ")
    text = Replace(text, vbTab, " ")
    text = Trim$(text)
    Do While Len(text) > 0
        ch = Left$(text, 1)
        If ch <> vbCr And ch <> vbLf And ch <> " " Then Exit Do
        text = Mid$(text, 2)
    Loop
    Do While Len(text) > 0
        ch = Right$(text, 1)
        If ch <> vbCr And ch <> vbLf And ch <> " " Then Exit Do
        text = Left$(text, Len(text) - 1)
    Loop
    NormalizeCellText = text
End Function

Private Function IsConvertibleNumber(ByVal text As String) As Boolean
    Dim body As String
    body = text
    If Left$(body, 1) = "+" Or Left$(body, 1) = "-" Then body = Mid$(body, 2)
    If Len(body) = 0 Then Exit Function
    If Left$(body, 1) = "0" And body <> "0" And Left$(body, 2) <> "0." Then Exit Function
    IsConvertibleNumber = IsNumeric(text)
End Function
