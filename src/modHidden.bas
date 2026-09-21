Attribute VB_Name = "modHidden"
Option Explicit

' 非表示のシート、行、列、白い文字、極小の文字、非表示の名前を探す。
' このモジュールだけでインポートして使えます。
' 非表示行と非表示列は、値があるときだけ報告する。
' 白文字は、文字色が白に近いか、塗りつぶしと同じ色のとき。
' 極小の文字は、サイズが 1 または 2。

Private Const MaxStyleHits As Long = 2000
Private mHits As Collection
Private mStyleHits As Long
Private mTruncated As Boolean

' 【FindHiddenContent】見えないまま残っているデータの一覧を二次元配列で返す。
'   列は 種別, シート, 場所, 内容。見つからなければ Empty。
'   種別は 非表示シート, 非常に非表示, 非表示行, 非表示列, 白文字, 同色文字, 極小文字, 非表示の名前, 参照先。
' 使用例:
'   hits = FindHiddenContent(ActiveWorkbook)
' 解説: 前面のブックで、隠れているシートや値のある非表示行、白いまま残った文字、非表示の名前を探す。非表示シートの中の行までは展開しない。
Public Function FindHiddenContent(Optional ByVal wb As Workbook) As Variant
    Dim ws As Worksheet
    Dim chartSheet As Chart
    On Error GoTo EH
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise 5, "FindHiddenContent", "調べるブックがありません。"
    Set mHits = New Collection
    mStyleHits = 0
    mTruncated = False
    For Each ws In wb.Worksheets
        ScanSheetVisibility ws.Name, ws.Visible, UsedAddress(ws)
        If ws.Visible = xlSheetVisible Then ScanSheetDetail ws
    Next ws
    For Each chartSheet In wb.Charts
        ScanSheetVisibility chartSheet.Name, chartSheet.Visible, "グラフシート"
    Next chartSheet
    ScanNames wb
    FindHiddenContent = HitsToArray()
    Set mHits = Nothing
    Exit Function
EH:
    Set mHits = Nothing
    Err.Raise Err.Number, "FindHiddenContent", Err.Description
End Function

' 【WriteHiddenContent】一覧を destination へ書き、見出しを含む行数を返す。
'   該当が無ければ 2 行目に「（なし）」。
' 使用例:
'   rows = WriteHiddenContent(Worksheets("結果").Range("A1"), ActiveWorkbook)
' 解説: 前面のブックの非表示データを、「結果」の A1 から種別・シート・場所・内容で書く。
Public Function WriteHiddenContent(ByVal destination As Range, Optional ByVal wb As Workbook) As Long
    Dim hits As Variant
    Dim labels As Variant
    Dim i As Long
    Dim rows As Long
    If destination Is Nothing Then Err.Raise 5, "WriteHiddenContent", "書き込み先が指定されていません。"
    hits = FindHiddenContent(wb)
    labels = Array("種別", "シート", "場所", "内容")
    For i = 0 To 3
        destination.Cells(1, i + 1).Value = labels(i)
    Next i
    destination.Resize(1, 4).Font.Bold = True
    If IsEmpty(hits) Then
        destination.Cells(2, 1).Value = "（なし）"
        WriteHiddenContent = 2
        Exit Function
    End If
    rows = UBound(hits, 1)
    WriteHitRows destination, hits, rows
    WriteHiddenContent = rows + 1
End Function

' 【Run_FindHiddenContent】マクロ一覧用。前面のブックを調べ、「非表示確認」シートへ書く。
'   実行前に、そのシートのセルは空にする。
' 使用例:
'   提出前のブックを前面にして実行する。
' 解説: 隠れているシート、値のある非表示行と列、白文字、極小文字、非表示の名前を一覧にし、件数を表示する。
Public Sub Run_FindHiddenContent()
    Dim wb As Workbook
    Dim report As Worksheet
    Dim written As Long
    Dim found As Long
    Dim prevScreen As Boolean
    Dim screenChanged As Boolean
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "調べるブックがありません。", vbExclamation, "非表示確認"
        Exit Sub
    End If
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenChanged = True
    Set report = EnsureSheet(wb, "非表示確認")
    report.Cells.Clear
    written = WriteHiddenContent(report.Range("A1"), wb)
    report.Columns("A:D").AutoFit
    report.Activate
    report.Range("A1").Select
    Application.ScreenUpdating = prevScreen
    If report.Range("A2").Value = "（なし）" Then
        MsgBox "見えないまま残っているデータは見つかりませんでした。", vbInformation, "非表示確認"
    Else
        found = written - 1
        If report.Cells(written, 1).Value = "（省略）" Then found = found - 1
        MsgBox CStr(found) & " 件見つかりました。結果は「非表示確認」シートにあります。", vbExclamation, "非表示確認"
    End If
    Exit Sub
EH:
    If screenChanged Then Application.ScreenUpdating = prevScreen
    MsgBox Err.Description, vbExclamation, "非表示確認"
End Sub

Private Sub ScanSheetVisibility(ByVal sheetName As String, ByVal visibleState As XlSheetVisibility, ByVal detail As String)
    Select Case visibleState
        Case xlSheetHidden
            AddHit "非表示シート", sheetName, "", detail
        Case xlSheetVeryHidden
            AddHit "非常に非表示", sheetName, "", detail
    End Select
End Sub

Private Sub ScanSheetDetail(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Sub
    ScanHiddenRows ws, lastRow, lastCol
    ScanHiddenColumns ws, lastRow, lastCol
    ScanFonts ws, lastRow, lastCol
End Sub

Private Sub ScanHiddenRows(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim rowIndex As Long
    Dim startRow As Long
    startRow = 0
    For rowIndex = 1 To lastRow
        If ws.Rows(rowIndex).Hidden Then
            If startRow = 0 Then startRow = rowIndex
        ElseIf startRow > 0 Then
            ReportHiddenBlock ws, startRow, rowIndex - 1, 1, lastCol, True
            startRow = 0
        End If
    Next rowIndex
    If startRow > 0 Then ReportHiddenBlock ws, startRow, lastRow, 1, lastCol, True
End Sub

Private Sub ScanHiddenColumns(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim colIndex As Long
    Dim startCol As Long
    startCol = 0
    For colIndex = 1 To lastCol
        If ws.Columns(colIndex).Hidden Then
            If startCol = 0 Then startCol = colIndex
        ElseIf startCol > 0 Then
            ReportHiddenBlock ws, 1, lastRow, startCol, colIndex - 1, False
            startCol = 0
        End If
    Next colIndex
    If startCol > 0 Then ReportHiddenBlock ws, 1, lastRow, startCol, lastCol, False
End Sub

Private Sub ReportHiddenBlock(ByVal ws As Worksheet, ByVal row1 As Long, ByVal row2 As Long, _
    ByVal col1 As Long, ByVal col2 As Long, ByVal isRow As Boolean)

    Dim sample As String
    Dim place As String
    sample = FirstContent(ws.Range(ws.Cells(row1, col1), ws.Cells(row2, col2)).Value)
    If Len(sample) = 0 Then Exit Sub
    If isRow Then
        place = SpanText(row1, row2)
        AddHit "非表示行", ws.Name, place, sample
    Else
        place = ColumnLetterOf(col1)
        If col2 <> col1 Then place = place & ":" & ColumnLetterOf(col2)
        AddHit "非表示列", ws.Name, place, sample
    End If
End Sub

Private Sub ScanFonts(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim values As Variant
    Dim target As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim cell As Range
    Set target = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    If target.Cells.CountLarge = 1 Then
        InspectFont ws, target, target.Value
        Exit Sub
    End If
    values = target.Value
    For rowIndex = 1 To lastRow
        If ws.Rows(rowIndex).Hidden Then GoTo NextRow
        For colIndex = 1 To lastCol
            If mTruncated Then Exit Sub
            If ws.Columns(colIndex).Hidden Then GoTo NextCol
            If HasContent(values(rowIndex, colIndex)) Then
                Set cell = ws.Cells(rowIndex, colIndex)
                InspectFont ws, cell, values(rowIndex, colIndex)
            End If
NextCol:
        Next colIndex
NextRow:
    Next rowIndex
End Sub

Private Sub InspectFont(ByVal ws As Worksheet, ByVal cell As Range, ByVal value As Variant)
    Dim fontColor As Long
    Dim fillColor As Long
    Dim fontSize As Double
    Dim hasFont As Boolean
    Dim nearWhite As Boolean
    Dim sameColor As Boolean
    Dim place As String
    If mStyleHits >= MaxStyleHits Then
        mTruncated = True
        Exit Sub
    End If
    On Error Resume Next
    fontColor = cell.Font.Color
    hasFont = (Err.Number = 0)
    Err.Clear
    fontSize = cell.Font.Size
    fillColor = cell.Interior.Color
    If cell.Interior.Pattern <> xlSolid Then fillColor = -1
    On Error GoTo 0
    If Not hasFont Then Exit Sub
    place = cell.Address(False, False)
    nearWhite = IsNearWhite(fontColor)
    sameColor = (fillColor >= 0 And fontColor = fillColor And Not nearWhite)
    If nearWhite Then
        AddStyleHit "白文字", ws.Name, place, Snippet(ContentText(value))
    ElseIf sameColor Then
        AddStyleHit "同色文字", ws.Name, place, Snippet(ContentText(value))
    End If
    If fontSize > 0 And fontSize <= 2 Then
        AddStyleHit "極小文字", ws.Name, place, Snippet(ContentText(value))
    End If
End Sub

Private Sub ScanNames(ByVal wb As Workbook)
    Dim definedName As Name
    Dim target As Range
    Dim sheetName As String
    Dim refersTo As String
    Dim rowState As String
    Dim colState As String
    Dim isVisible As Boolean
    Dim visibleKnown As Boolean
    For Each definedName In wb.Names
        refersTo = ""
        sheetName = ""
        isVisible = True
        visibleKnown = False
        Set target = Nothing
        On Error Resume Next
        refersTo = definedName.RefersTo
        Set target = definedName.RefersToRange
        Err.Clear
        isVisible = definedName.Visible
        visibleKnown = (Err.Number = 0)
        On Error GoTo 0
        If visibleKnown And Not isVisible Then
            If Not target Is Nothing Then sheetName = target.Worksheet.Name
            AddHit "非表示の名前", sheetName, definedName.Name, refersTo
        End If
        If Not target Is Nothing Then
            sheetName = target.Worksheet.Name
            If target.Worksheet.Visible <> xlSheetVisible Then
                AddHit "参照先", sheetName, definedName.Name, "非表示シートを参照"
            ElseIf target.Rows.Count <= 500 And target.Columns.Count <= 100 Then
                rowState = HiddenState(target.EntireRow)
                colState = HiddenState(target.EntireColumn)
                If rowState = "hidden" Then AddHit "参照先", sheetName, definedName.Name, "非表示行を参照"
                If colState = "hidden" Then AddHit "参照先", sheetName, definedName.Name, "非表示列を参照"
            End If
        End If
    Next definedName
End Sub

Private Function HiddenState(ByVal target As Range) As String
    Dim state As Variant
    On Error Resume Next
    state = target.Hidden
    On Error GoTo 0
    If VarType(state) = vbBoolean Then
        If state = True Then
            HiddenState = "hidden"
        Else
            HiddenState = "visible"
        End If
    Else
        HiddenState = "mixed"
    End If
End Function

Private Sub AddStyleHit(ByVal kind As String, ByVal sheetName As String, ByVal place As String, ByVal content As String)
    If mStyleHits >= MaxStyleHits Then
        mTruncated = True
        Exit Sub
    End If
    mStyleHits = mStyleHits + 1
    AddHit kind, sheetName, place, content
End Sub

Private Sub AddHit(ByVal kind As String, ByVal sheetName As String, ByVal place As String, ByVal content As String)
    If mHits Is Nothing Then Exit Sub
    mHits.Add Array(kind, sheetName, place, content)
End Sub

Private Function FirstContent(ByVal values As Variant) As String
    Dim rowIndex As Long
    Dim colIndex As Long
    If Not IsArray(values) Then
        If HasContent(values) Then FirstContent = Snippet(ContentText(values))
        Exit Function
    End If
    For rowIndex = LBound(values, 1) To UBound(values, 1)
        For colIndex = LBound(values, 2) To UBound(values, 2)
            If HasContent(values(rowIndex, colIndex)) Then
                FirstContent = Snippet(ContentText(values(rowIndex, colIndex)))
                Exit Function
            End If
        Next colIndex
    Next rowIndex
End Function

Private Function HasContent(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarType(value) = vbString Then
        HasContent = (Len(value) > 0)
    Else
        HasContent = True
    End If
End Function

Private Function ContentText(ByVal value As Variant) As String
    If IsError(value) Then
        ContentText = "エラー"
    Else
        ContentText = CStr(value)
    End If
End Function

Private Function Snippet(ByVal text As String) As String
    text = Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " ")
    If Len(text) > 80 Then text = Left$(text, 80) & "..."
    Snippet = text
End Function

Private Function IsNearWhite(ByVal color As Long) As Boolean
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long
    If color < 0 Then Exit Function
    redPart = color Mod 256
    greenPart = (color \ 256) Mod 256
    bluePart = (color \ 65536) Mod 256
    IsNearWhite = (redPart >= 250 And greenPart >= 250 And bluePart >= 250)
End Function

Private Function SpanText(ByVal startIndex As Long, ByVal endIndex As Long) As String
    If startIndex = endIndex Then
        SpanText = CStr(startIndex)
    Else
        SpanText = CStr(startIndex) & ":" & CStr(endIndex)
    End If
End Function

Private Function UsedAddress(ByVal ws As Worksheet) As String
    Dim rowIndex As Long
    Dim colIndex As Long
    rowIndex = LastUsedRow(ws)
    colIndex = LastUsedColumn(ws)
    If rowIndex = 0 Or colIndex = 0 Then
        UsedAddress = "（データなし）"
    Else
        UsedAddress = "A1:" & ColumnLetterOf(colIndex) & CStr(rowIndex)
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not lastCell Is Nothing Then LastUsedRow = lastCell.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not lastCell Is Nothing Then LastUsedColumn = lastCell.Column
End Function

Private Function HitsToArray() As Variant
    Dim data() As Variant
    Dim index As Long
    Dim item As Variant
    Dim total As Long
    If mHits Is Nothing Or mHits.Count = 0 Then Exit Function
    total = mHits.Count
    If mTruncated Then total = total + 1
    ReDim data(1 To total, 1 To 4)
    For index = 1 To mHits.Count
        item = mHits(index)
        data(index, 1) = item(0)
        data(index, 2) = item(1)
        data(index, 3) = item(2)
        data(index, 4) = item(3)
    Next index
    If mTruncated Then
        data(total, 1) = "（省略）"
        data(total, 4) = "白文字と極小文字は先頭 " & CStr(MaxStyleHits) & " 件までです。"
    End If
    HitsToArray = data
End Function

Private Sub WriteHitRows(ByVal destination As Range, ByVal hits As Variant, ByVal rows As Long)
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim text As String
    For rowIndex = 1 To rows
        For colIndex = 1 To 4
            If IsEmpty(hits(rowIndex, colIndex)) Then
                text = ""
            Else
                text = CStr(hits(rowIndex, colIndex))
            End If
            If Len(text) > 0 Then
                If Left$(text, 1) = "=" Or Left$(text, 1) = "'" Then text = "'" & text
            End If
            destination.Cells(rowIndex + 1, colIndex).Value = text
        Next colIndex
    Next rowIndex
End Sub

Private Function ColumnLetterOf(ByVal columnIndex As Long) As String
    Dim n As Long
    Dim remainder As Long
    n = columnIndex
    Do While n > 0
        remainder = (n - 1) Mod 26
        ColumnLetterOf = Chr$(65 + remainder) & ColumnLetterOf
        n = (n - 1) \ 26
    Loop
End Function

Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function
