Attribute VB_Name = "modForbidden"
Option Explicit

' 禁止ワードがブック内のセルとオブジェクトに含まれるかを調べる。
' このモジュールだけでインポートして使えます。
' 対象は、セルの値、数式、メモ、新しいコメント、図形の文字と代替テキスト、
' グラフのタイトル・軸・系列名・データラベル（各系列 2000 点まで）、ハイパーリンク。
' ヘッダー、フッター、VBA のコードは対象外。
' 禁止ワードにセル範囲を渡したとき、その範囲が調べるブック上にあれば、そのセル自身は該当にしない。

Private Const MaxHits As Long = 5000
Private Const MaxWords As Long = 2000
Private Const MaxLabelPoints As Long = 2000
Private Const ReportSheetName As String = "禁止ワード確認"

Private mWords() As String
Private mHits As Collection
Private mSkip As Object
Private mWb As Workbook
Private mSheet As String
Private mCompare As VbCompareMethod
Private mTruncated As Boolean

' 【FindForbiddenWords】禁止ワードと照合し、見つかった箇所を二次元配列で返す。
'   words はセル範囲、1 件の文字列、配列、Collection のいずれか。空欄は無視する。
'   列は 種別, シート, 場所, 禁止ワード, 内容。見つからなければ Empty。
'   既定では大文字小文字を区別せず、部分一致で探す。5000 件を超える分は省略する。
'   非表示と非常に非表示のシートも調べる。
' 使用例:
'   hits = FindForbiddenWords(Worksheets("設定").Range("A2:A30"), ActiveWorkbook)
'   hits = FindForbiddenWords(Array("社外秘", "パスワード"))
' 解説: 1 行目は「設定」の A2:A30 を禁止ワードとして、前面のブックを調べる。この範囲自体は該当にしない。2 行目は配列で渡した 2 語を、同じように調べる。
Public Function FindForbiddenWords(ByVal words As Variant, Optional ByVal wb As Workbook, _
    Optional ByVal matchCase As Boolean = False) As Variant

    Dim ws As Worksheet
    Dim chartSheet As Chart
    On Error GoTo EH
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise 5, "FindForbiddenWords", "調べるブックがありません。"
    BeginScan words, wb, matchCase
    For Each ws In wb.Worksheets
        mSheet = ws.Name
        ScanWorksheetCells ws
        ScanShapes ws.Shapes
        ScanNotes ws
        ScanThreadedComments ws
        ScanLinkCollection ws.Hyperlinks, ws.Name
        If mTruncated Then Exit For
    Next ws
    If Not mTruncated Then
        For Each chartSheet In wb.Charts
            mSheet = chartSheet.Name
            ScanChart chartSheet, chartSheet.Name
            ScanShapes chartSheet.Shapes
            ScanLinkCollection chartSheet.Hyperlinks, chartSheet.Name
            If mTruncated Then Exit For
        Next chartSheet
    End If
    FindForbiddenWords = HitsToArray()
    ResetScan
    Exit Function
EH:
    ResetScan
    Err.Raise Err.Number, "FindForbiddenWords", Err.Description
End Function

' 【WriteForbiddenWords】照合結果を destination へ書き、見出しを含む行数を返す。
'   該当が無ければ 2 行目に「（なし）」。= で始まる内容は、文字として書く。
'   前回の結果が同じシートに残っていると、それも該当になる。
' 使用例:
'   rows = WriteForbiddenWords(Worksheets("設定").Range("A2:A30"), Worksheets("結果").Range("A1"))
' 解説: 「設定」の一覧で前面のブックを調べ、「結果」の A1 から種別・シート・場所・禁止ワード・内容を書く。戻り値は見出しを含む行数。
Public Function WriteForbiddenWords(ByVal words As Variant, ByVal destination As Range, _
    Optional ByVal wb As Workbook, Optional ByVal matchCase As Boolean = False) As Long

    Dim hits As Variant
    Dim rows As Long
    Dim labels As Variant
    Dim i As Long
    If destination Is Nothing Then Err.Raise 5, "WriteForbiddenWords", "書き込み先が指定されていません。"
    hits = FindForbiddenWords(words, wb, matchCase)
    labels = Array("種別", "シート", "場所", "禁止ワード", "内容")
    For i = 0 To 4
        destination.Cells(1, i + 1).Value = labels(i)
    Next i
    destination.Resize(1, 5).Font.Bold = True
    If IsEmpty(hits) Then
        destination.Cells(2, 1).Value = "（なし）"
        WriteForbiddenWords = 2
        Exit Function
    End If
    rows = UBound(hits, 1)
    WriteHitRows destination, hits, rows
    WriteForbiddenWords = rows + 1
End Function

' 【Run_FindForbiddenWords】マクロ一覧用。禁止ワードの範囲を選ばせ、前面のブックを調べる。
'   結果は「禁止ワード確認」シートへ書く。このシート上の範囲は一覧にできない。
'   一覧に使ったセル自身は該当にしない。検索の前に、結果シートのセルは空にする。
' 使用例:
'   禁止ワードを縦に並べた範囲を用意し、調べたいブックを前面にして実行する。
' 解説: 範囲を選んで OK すると、セル、図形、グラフ、メモ、ハイパーリンクを調べ、該当件数を表示する。キャンセルすると何もしない。
Public Sub Run_FindForbiddenWords()
    Dim picked As Range
    Dim report As Worksheet
    Dim wb As Workbook
    Dim written As Long
    Dim found As Long
    Dim prevScreen As Boolean
    Dim screenChanged As Boolean
    On Error Resume Next
    Set picked = Application.InputBox( _
        "禁止ワードが入ったセル範囲を選択してください。" & vbCrLf & _
        "前面のブックのセル、図形、グラフ、メモ、ハイパーリンクを調べます。", _
        "禁止ワード", Type:=8)
    If Err.Number <> 0 Or picked Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "調べるブックがありません。", vbExclamation, "禁止ワード確認"
        Exit Sub
    End If
    If StrComp(picked.Worksheet.Name, ReportSheetName, vbTextCompare) = 0 _
        And picked.Worksheet.Parent Is wb Then
        MsgBox "禁止ワードの一覧は「" & ReportSheetName & "」以外のシートに置いてください。", vbExclamation, "禁止ワード確認"
        Exit Sub
    End If
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenChanged = True
    Set report = EnsureReportSheet(wb)
    report.Cells.Clear
    written = WriteForbiddenWords(picked, report.Range("A1"), wb)
    report.Columns("A:E").AutoFit
    report.Activate
    report.Range("A1").Select
    Application.ScreenUpdating = prevScreen
    If report.Range("A2").Value = "（なし）" Then
        MsgBox "禁止ワードは見つかりませんでした。", vbInformation, "禁止ワード確認"
    Else
        found = written - 1
        If report.Cells(written, 1).Value = "（省略）" Then found = found - 1
        MsgBox CStr(found) & " 件見つかりました。結果は「" & ReportSheetName & "」シートにあります。", vbExclamation, "禁止ワード確認"
    End If
    Exit Sub
EH:
    If screenChanged Then Application.ScreenUpdating = prevScreen
    MsgBox Err.Description, vbExclamation, "禁止ワード確認"
End Sub

Private Sub BeginScan(ByVal words As Variant, ByVal wb As Workbook, ByVal matchCase As Boolean)
    Set mHits = New Collection
    Set mSkip = CreateObject("Scripting.Dictionary")
    Set mWb = wb
    mTruncated = False
    mSheet = ""
    If matchCase Then
        mCompare = vbBinaryCompare
    Else
        mCompare = vbTextCompare
    End If
    mWords = NormalizeWords(words)
    If TypeName(words) = "Range" Then AddSkipRange words
End Sub

Private Sub ResetScan()
    Set mHits = Nothing
    Set mSkip = Nothing
    Set mWb = Nothing
    mTruncated = False
    mSheet = ""
    Erase mWords
End Sub

Private Function NormalizeWords(ByVal words As Variant) As String()
    Dim dict As Object
    Dim key As Variant
    Dim result() As String
    Dim index As Long
    Set dict = CreateObject("Scripting.Dictionary")
    If mCompare = vbBinaryCompare Then
        dict.CompareMode = 0
    Else
        dict.CompareMode = 1
    End If
    AppendWords dict, words
    If dict.Count = 0 Then Err.Raise 5, "FindForbiddenWords", "禁止ワードがありません。空欄ではないセルを指定してください。"
    If dict.Count > MaxWords Then Err.Raise 5, "FindForbiddenWords", "禁止ワードは " & CStr(MaxWords) & " 件までです。"
    ReDim result(1 To dict.Count)
    For Each key In dict.Keys
        index = index + 1
        result(index) = CStr(key)
    Next key
    NormalizeWords = result
End Function

Private Sub AppendWords(ByVal dict As Object, ByVal words As Variant)
    Dim cell As Range
    Dim item As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    If IsObject(words) Then
        If TypeName(words) = "Range" Then
            If words.Cells.CountLarge > MaxWords Then
                Err.Raise 5, "FindForbiddenWords", "禁止ワードは " & CStr(MaxWords) & " 件までです。"
            End If
            For Each cell In words.Cells
                AddWord dict, cell.Value
            Next cell
            Exit Sub
        End If
        If TypeName(words) = "Collection" Then
            For Each item In words
                AddWord dict, item
                If dict.Count > MaxWords Then Exit For
            Next item
            Exit Sub
        End If
    End If
    If VarType(words) = vbString Then
        AddWord dict, words
        Exit Sub
    End If
    If IsArray(words) Then
        If DimensionCount(words) <= 1 Then
            For rowIndex = LBound(words) To UBound(words)
                AddWord dict, words(rowIndex)
            Next rowIndex
        Else
            For rowIndex = LBound(words, 1) To UBound(words, 1)
                For colIndex = LBound(words, 2) To UBound(words, 2)
                    AddWord dict, words(rowIndex, colIndex)
                Next colIndex
            Next rowIndex
        End If
        Exit Sub
    End If
    AddWord dict, words
End Sub

Private Sub AddWord(ByVal dict As Object, ByVal value As Variant)
    Dim text As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Or IsObject(value) Then Exit Sub
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Sub
    If Len(text) > 200 Then Err.Raise 5, "FindForbiddenWords", "禁止ワードが長すぎます。"
    If Not dict.Exists(text) Then dict.Add text, True
End Sub

Private Sub AddSkipRange(ByVal words As Range)
    Dim area As Range
    Dim cell As Range
    For Each area In words.Areas
        If area.Worksheet.Parent Is mWb Then
            For Each cell In area.Cells
                mSkip(SkipKey(area.Worksheet.Name, cell.Row, cell.Column)) = True
            Next cell
        End If
    Next area
End Sub

Private Function SkipKey(ByVal sheetName As String, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    SkipKey = sheetName & Chr$(1) & CStr(rowIndex) & Chr$(1) & CStr(columnIndex)
End Function

Private Function CellSkipped(ByVal sheetName As String, ByVal rowIndex As Long, ByVal columnIndex As Long) As Boolean
    If mSkip Is Nothing Then Exit Function
    CellSkipped = mSkip.Exists(SkipKey(sheetName, rowIndex, columnIndex))
End Function

Private Sub ScanWorksheetCells(ByVal ws As Worksheet)
    Dim target As Range
    Dim values As Variant
    Dim formulas As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim place As String
    Set target = UsedDataRange(ws)
    If target Is Nothing Then Exit Sub
    If target.Cells.CountLarge = 1 Then
        If Not CellSkipped(ws.Name, target.Row, target.Column) Then
            ReportCellTexts target.Address(False, False), CellText(target.Value), CellFormula(target.Formula)
        End If
        Exit Sub
    End If
    values = target.Value
    formulas = target.Formula
    For rowIndex = 1 To UBound(values, 1)
        For colIndex = 1 To UBound(values, 2)
            If mTruncated Then Exit Sub
            If Not CellSkipped(ws.Name, target.Row + rowIndex - 1, target.Column + colIndex - 1) Then
                place = ColumnLetterOf(target.Column + colIndex - 1) & CStr(target.Row + rowIndex - 1)
                ReportCellTexts place, CellText(values(rowIndex, colIndex)), CellFormula(formulas(rowIndex, colIndex))
            End If
        Next colIndex
    Next rowIndex
End Sub

Private Sub ReportCellTexts(ByVal place As String, ByVal valueText As String, ByVal formulaText As String)
    Dim index As Long
    If Len(valueText) = 0 And Len(formulaText) = 0 Then Exit Sub
    For index = 1 To UBound(mWords)
        If mTruncated Then Exit Sub
        If Len(valueText) > 0 And ContainsWord(valueText, mWords(index)) Then
            AddHit "セル", mSheet, place, mWords(index), valueText
        ElseIf Len(formulaText) > 0 And ContainsWord(formulaText, mWords(index)) Then
            AddHit "数式", mSheet, place, mWords(index), formulaText
        End If
    Next index
End Sub

Private Sub ScanShapes(ByVal shapes As Shapes)
    Dim shp As Shape
    For Each shp In shapes
        If mTruncated Then Exit Sub
        ScanShape shp, 0
    Next shp
End Sub

Private Sub ScanShape(ByVal shp As Shape, ByVal depth As Long)
    Dim index As Long
    Dim childCount As Long
    Dim caption As String
    Dim altText As String
    Dim chartObject As Chart
    Dim isGroup As Boolean
    If depth > 15 Or mTruncated Then Exit Sub
    On Error Resume Next
    isGroup = (shp.Type = msoGroup)
    If isGroup Then childCount = shp.GroupItems.Count
    If shp.HasChart Then Set chartObject = shp.Chart
    caption = ShapeCaption(shp)
    altText = shp.AlternativeText
    On Error GoTo 0
    If isGroup Then
        For index = 1 To childCount
            ScanShape shp.GroupItems(index), depth + 1
        Next index
    End If
    If Not chartObject Is Nothing Then ScanChart chartObject, shp.Name
    ScanText caption, "図形", mSheet, shp.Name
    If Len(altText) > 0 Then ScanText altText, "図形", mSheet, shp.Name & " 代替テキスト"
    ScanSmartArt shp
End Sub

Private Sub ScanSmartArt(ByVal shp As Shape)
    Dim nodes As Object
    Dim node As Object
    Dim text As String
    On Error Resume Next
    If shp.HasSmartArt <> True Then Exit Sub
    Set nodes = shp.SmartArt.AllNodes
    On Error GoTo 0
    If nodes Is Nothing Then Exit Sub
    For Each node In nodes
        If mTruncated Then Exit Sub
        text = ""
        On Error Resume Next
        text = CStr(node.TextFrame2.TextRange.Text)
        On Error GoTo 0
        ScanText text, "図形", mSheet, shp.Name
    Next node
End Sub

Private Function ShapeCaption(ByVal shp As Shape) As String
    On Error Resume Next
    If shp.HasTextFrame Then
        If shp.TextFrame2.HasText Then ShapeCaption = shp.TextFrame2.TextRange.Text
        If Len(ShapeCaption) = 0 And shp.TextFrame.HasText Then
            ShapeCaption = shp.TextFrame.Characters.Text
        End If
    End If
    If Len(ShapeCaption) = 0 Then ShapeCaption = CStr(shp.OLEFormat.Object.Caption)
    If Len(ShapeCaption) = 0 Then ShapeCaption = CStr(shp.OLEFormat.Object.Text)
    On Error GoTo 0
End Function

Private Sub ScanChart(ByVal chartObject As Chart, ByVal chartName As String)
    Dim bag As Collection
    Dim item As Variant
    Set bag = New Collection
    CollectChartTexts chartObject, chartName, bag
    For Each item In bag
        If mTruncated Then Exit Sub
        ScanText CStr(item(1)), "グラフ", mSheet, CStr(item(0))
    Next item
End Sub

Private Sub CollectChartTexts(ByVal chartObject As Chart, ByVal chartName As String, ByVal bag As Collection)
    Dim axisItem As Axis
    Dim seriesItem As Series
    Dim pointIndex As Long
    Dim pointCount As Long
    Dim text As String
    On Error Resume Next
    text = ""
    If chartObject.HasTitle Then text = CStr(chartObject.ChartTitle.Text)
    If Len(text) > 0 Then bag.Add Array(chartName & " タイトル", text)
    For Each axisItem In chartObject.Axes
        text = ""
        If axisItem.HasTitle Then text = CStr(axisItem.AxisTitle.Text)
        If Len(text) > 0 Then bag.Add Array(chartName & " 軸", text)
    Next axisItem
    For Each seriesItem In chartObject.SeriesCollection
        text = ""
        text = CStr(seriesItem.Name)
        If Len(text) > 0 Then bag.Add Array(chartName & " 系列", text)
        pointCount = 0
        If seriesItem.HasDataLabels Then pointCount = seriesItem.Points.Count
        If pointCount > MaxLabelPoints Then pointCount = MaxLabelPoints
        For pointIndex = 1 To pointCount
            text = ""
            If seriesItem.Points(pointIndex).HasDataLabel Then
                text = CStr(seriesItem.Points(pointIndex).DataLabel.Text)
            End If
            If Len(text) > 0 Then bag.Add Array(chartName & " データラベル", text)
        Next pointIndex
    Next seriesItem
    On Error GoTo 0
End Sub

Private Sub ScanNotes(ByVal ws As Worksheet)
    Dim noteItem As Comment
    Dim place As String
    For Each noteItem In ws.Comments
        If mTruncated Then Exit Sub
        If Not CellSkipped(ws.Name, noteItem.Parent.Row, noteItem.Parent.Column) Then
            place = noteItem.Parent.Address(False, False)
            ScanText noteItem.Text, "メモ", ws.Name, place
        End If
    Next noteItem
End Sub

Private Sub ScanThreadedComments(ByVal ws As Worksheet)
    Dim threads As Object
    Dim threadItem As Object
    Dim reply As Object
    Dim replies As Object
    Dim text As String
    Dim place As String
    Dim rowIndex As Long
    Dim colIndex As Long
    On Error Resume Next
    Set threads = CallByName(ws, "CommentsThreaded", VbGet)
    On Error GoTo 0
    If threads Is Nothing Then Exit Sub
    For Each threadItem In threads
        If mTruncated Then Exit Sub
        text = ""
        place = ""
        rowIndex = 0
        colIndex = 0
        On Error Resume Next
        text = CStr(threadItem.Text)
        place = threadItem.Parent.Address(False, False)
        rowIndex = threadItem.Parent.Row
        colIndex = threadItem.Parent.Column
        Set replies = threadItem.Replies
        On Error GoTo 0
        If rowIndex > 0 And Not CellSkipped(ws.Name, rowIndex, colIndex) Then
            ScanText text, "コメント", ws.Name, place
            If Not replies Is Nothing Then
                For Each reply In replies
                    text = ""
                    On Error Resume Next
                    text = CStr(reply.Text)
                    On Error GoTo 0
                    ScanText text, "コメント", ws.Name, place
                Next reply
            End If
        End If
    Next threadItem
End Sub

Private Sub ScanLinkCollection(ByVal links As Hyperlinks, ByVal sheetName As String)
    Dim link As Hyperlink
    Dim place As String
    For Each link In links
        If mTruncated Then Exit Sub
        place = ""
        If link.Type = msoHyperlinkRange Then
            If Not CellSkipped(sheetName, link.Range.Row, link.Range.Column) Then
                place = link.Range.Address(False, False)
            End If
        Else
            On Error Resume Next
            place = link.Shape.Name
            On Error GoTo 0
            If Len(place) = 0 Then place = "図形"
        End If
        If Len(place) > 0 Then
            ScanText link.Address, "ハイパーリンク", sheetName, place & " アドレス"
            ScanText link.SubAddress, "ハイパーリンク", sheetName, place & " サブアドレス"
            ScanText link.ScreenTip, "ハイパーリンク", sheetName, place & " ポップヒント"
            ScanText link.TextToDisplay, "ハイパーリンク", sheetName, place & " 表示文字"
        End If
    Next link
End Sub

Private Sub WriteHitRows(ByVal destination As Range, ByVal hits As Variant, ByVal rows As Long)
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim text As String
    For rowIndex = 1 To rows
        For colIndex = 1 To 5
            text = CStr(hits(rowIndex, colIndex))
            If Len(text) > 0 Then
                If Left$(text, 1) = "=" Or Left$(text, 1) = "'" Then text = "'" & text
            End If
            destination.Cells(rowIndex + 1, colIndex).Value = text
        Next colIndex
    Next rowIndex
End Sub

Private Sub ScanText(ByVal text As String, ByVal kind As String, ByVal sheetName As String, ByVal place As String)
    Dim index As Long
    If Len(text) = 0 Or mTruncated Then Exit Sub
    For index = 1 To UBound(mWords)
        If mTruncated Then Exit Sub
        If ContainsWord(text, mWords(index)) Then
            AddHit kind, sheetName, place, mWords(index), text
        End If
    Next index
End Sub

Private Sub AddHit(ByVal kind As String, ByVal sheetName As String, ByVal place As String, _
    ByVal word As String, ByVal content As String)

    If mHits.Count >= MaxHits Then
        mTruncated = True
        Exit Sub
    End If
    mHits.Add Array(kind, sheetName, place, word, Snippet(content, word))
End Sub

Private Function HitsToArray() As Variant
    Dim data() As Variant
    Dim index As Long
    Dim item As Variant
    Dim total As Long
    If mHits Is Nothing Then Exit Function
    If mHits.Count = 0 Then Exit Function
    total = mHits.Count
    If mTruncated Then total = total + 1
    ReDim data(1 To total, 1 To 5)
    For index = 1 To mHits.Count
        item = mHits(index)
        data(index, 1) = item(0)
        data(index, 2) = item(1)
        data(index, 3) = item(2)
        data(index, 4) = item(3)
        data(index, 5) = item(4)
    Next index
    If mTruncated Then
        data(total, 1) = "（省略）"
        data(total, 5) = "該当は先頭 " & CStr(MaxHits) & " 件までです。"
    End If
    HitsToArray = data
End Function

Private Function ContainsWord(ByVal text As String, ByVal word As String) As Boolean
    ContainsWord = (InStr(1, text, word, mCompare) > 0)
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarType(value) <> vbString Then Exit Function
    CellText = CStr(value)
End Function

Private Function CellFormula(ByVal formula As Variant) As String
    Dim text As String
    If IsError(formula) Or IsEmpty(formula) Then Exit Function
    text = CStr(formula)
    If Left$(text, 1) <> "=" Then Exit Function
    CellFormula = text
End Function

Private Function Snippet(ByVal content As String, ByVal word As String) As String
    Dim pos As Long
    Dim startAt As Long
    content = Replace(Replace(Replace(content, vbCr, " "), vbLf, " "), vbTab, " ")
    If Len(content) <= 160 Then
        Snippet = content
        Exit Function
    End If
    pos = InStr(1, content, word, mCompare)
    If pos < 1 Then pos = 1
    startAt = pos - 50
    If startAt < 1 Then startAt = 1
    Snippet = Mid$(content, startAt, 160)
    If startAt > 1 Then Snippet = "..." & Snippet
    If startAt + 159 < Len(content) Then Snippet = Snippet & "..."
End Function

Private Function UsedDataRange(ByVal ws As Worksheet) As Range
    Dim lastCell As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Exit Function
    rowIndex = lastCell.Row
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Exit Function
    colIndex = lastCell.Column
    Set UsedDataRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowIndex, colIndex))
End Function

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

Private Function DimensionCount(ByVal values As Variant) As Long
    Dim n As Long
    Dim lower As Long
    On Error GoTo Done
    Do
        n = n + 1
        lower = LBound(values, n)
    Loop
Done:
    DimensionCount = n - 1
End Function

Private Function EnsureReportSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set EnsureReportSheet = wb.Worksheets(ReportSheetName)
    On Error GoTo 0
    If EnsureReportSheet Is Nothing Then
        Set EnsureReportSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureReportSheet.Name = ReportSheetName
    End If
End Function
