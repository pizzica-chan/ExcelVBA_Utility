Attribute VB_Name = "modPrivacy"
Option Explicit

' メール、電話、郵便番号、カード番号、個人番号に見える文字列を探す。
' このモジュールだけでインポートして使えます。
' 対象は、文字列として入っているセルの値と数式、メモ、コメント、図形、グラフ、ハイパーリンク。
' 数値セルに入った連続数字は対象外。電話は区切りのある番号、郵便番号はハイフン付き。
' カード番号は Luhn、個人番号は検査用数字が合うものだけ。結果の番号とメールはマスクする。郵便番号はそのまま出す。

Private Const MaxHits As Long = 3000
Private mHits As Collection
Private mSeen As Object
Private mSheet As String
Private mValueEcho As String
Private mTruncated As Boolean

' 【FindPersonalData】個人情報に見える文字列の該当箇所を二次元配列で返す。
'   列は 種別, シート, 場所, 該当。見つからなければ Empty。
'   種別は メール, 電話, 郵便番号, カード番号, 個人番号。
'   非表示と非常に非表示のシートも調べる。3000 件を超える分は省略する。
' 使用例:
'   hits = FindPersonalData(ActiveWorkbook)
' 解説: 前面のブックから、メールアドレスや 090-1234-5678 のような電話、〒100-0001、検査数字の合うカード番号と個人番号を探す。該当列は番号を伏せた文字列。
Public Function FindPersonalData(Optional ByVal wb As Workbook) As Variant
    Dim ws As Worksheet
    Dim chartSheet As Chart
    On Error GoTo EH
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise 5, "FindPersonalData", "調べるブックがありません。"
    BeginScan
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
    FindPersonalData = HitsToArray()
    ResetScan
    Exit Function
EH:
    ResetScan
    Err.Raise Err.Number, "FindPersonalData", Err.Description
End Function

' 【WritePersonalData】照合結果を destination へ書き、見出しを含む行数を返す。
'   該当が無ければ 2 行目に「（なし）」。= で始まる内容は文字として書く。
' 使用例:
'   rows = WritePersonalData(Worksheets("結果").Range("A1"), ActiveWorkbook)
' 解説: 前面のブックを調べ、「結果」の A1 から種別・シート・場所・該当を書く。メールと番号はマスクされている。
Public Function WritePersonalData(ByVal destination As Range, Optional ByVal wb As Workbook) As Long
    WritePersonalData = WriteHits(destination, FindPersonalData(wb), "WritePersonalData")
End Function

' 【Run_FindPersonalData】マクロ一覧用。前面のブックを調べ、「個人情報確認」シートへ書く。
'   実行前に、そのシートのセルは空にする。
' 使用例:
'   提出前のブックを前面にして実行する。
' 解説: メール、電話、郵便番号、カード番号、個人番号に見える文字列の件数を表示し、場所の一覧を「個人情報確認」に書く。
Public Sub Run_FindPersonalData()
    RunReport "個人情報確認", "WritePersonalData"
End Sub

Private Sub RunReport(ByVal sheetName As String, ByVal source As String)
    Dim wb As Workbook
    Dim report As Worksheet
    Dim written As Long
    Dim found As Long
    Dim prevScreen As Boolean
    Dim screenChanged As Boolean
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "調べるブックがありません。", vbExclamation, sheetName
        Exit Sub
    End If
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenChanged = True
    Set report = EnsureSheet(wb, sheetName)
    report.Cells.Clear
    written = WritePersonalData(report.Range("A1"), wb)
    FinishReport report, written, sheetName
    Application.ScreenUpdating = prevScreen
    Exit Sub
EH:
    If screenChanged Then Application.ScreenUpdating = prevScreen
    MsgBox Err.Description, vbExclamation, source
End Sub

Private Sub FinishReport(ByVal report As Worksheet, ByVal written As Long, ByVal title As String)
    Dim found As Long
    report.Columns("A:D").AutoFit
    report.Activate
    report.Range("A1").Select
    Application.ScreenUpdating = True
    If report.Range("A2").Value = "（なし）" Then
        MsgBox "個人情報に見える文字列は見つかりませんでした。", vbInformation, title
    Else
        found = written - 1
        If report.Cells(written, 1).Value = "（省略）" Then found = found - 1
        MsgBox CStr(found) & " 件見つかりました。結果は「" & title & "」シートにあります。", vbExclamation, title
    End If
End Sub

Private Function WriteHits(ByVal destination As Range, ByVal hits As Variant, ByVal source As String) As Long
    Dim labels As Variant
    Dim i As Long
    Dim rows As Long
    If destination Is Nothing Then Err.Raise 5, source, "書き込み先が指定されていません。"
    labels = Array("種別", "シート", "場所", "該当")
    For i = 0 To 3
        destination.Cells(1, i + 1).Value = labels(i)
    Next i
    destination.Resize(1, 4).Font.Bold = True
    If IsEmpty(hits) Then
        destination.Cells(2, 1).Value = "（なし）"
        WriteHits = 2
        Exit Function
    End If
    rows = UBound(hits, 1)
    WriteHitRows destination, hits, rows
    WriteHits = rows + 1
End Function

Private Sub BeginScan()
    Set mHits = New Collection
    Set mSeen = CreateObject("Scripting.Dictionary")
    mTruncated = False
    mSheet = ""
    mValueEcho = ""
End Sub

Private Sub ResetScan()
    Set mHits = Nothing
    Set mSeen = Nothing
    mTruncated = False
    mSheet = ""
    mValueEcho = ""
End Sub

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
        place = target.Address(False, False)
        ConsiderCell place, CellText(target.Value), CellFormula(target.Formula)
        Exit Sub
    End If
    values = target.Value
    formulas = target.Formula
    For rowIndex = 1 To UBound(values, 1)
        For colIndex = 1 To UBound(values, 2)
            If mTruncated Then Exit Sub
            place = ColumnLetterOf(target.Column + colIndex - 1) & CStr(target.Row + rowIndex - 1)
            ConsiderCell place, CellText(values(rowIndex, colIndex)), CellFormula(formulas(rowIndex, colIndex))
        Next colIndex
    Next rowIndex
End Sub

Private Sub ConsiderCell(ByVal place As String, ByVal valueText As String, ByVal formulaText As String)
    mValueEcho = ""
    If Len(valueText) > 0 Then Consider "セル " & place, valueText
    If Len(formulaText) > 0 Then
        mValueEcho = NormalizeScan(valueText)
        Consider "数式 " & place, formulaText
        mValueEcho = ""
    End If
End Sub

Private Sub Consider(ByVal place As String, ByVal rawText As String)
    Dim text As String
    If mTruncated Or Len(rawText) = 0 Then Exit Sub
    If Not MayContain(rawText) Then Exit Sub
    text = NormalizeScan(rawText)
    FindEmails text, place
    FindLabeledNumbers text, place, "0\d{1,4}-\d{1,4}-\d{3,4}|\+81-?\d{1,4}-\d{1,4}-\d{3,4}", "電話", 10, 13
    FindPostals text, place
    FindLongNumbers text, place
End Sub

Private Sub FindEmails(ByVal text As String, ByVal place As String)
    Dim matches As Object
    Dim item As Object
    Dim email As String
    Set matches = MatchAll("[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", text)
    For Each item In matches
        If mTruncated Then Exit Sub
        email = CStr(item.Value)
        If Len(email) <= 254 Then AddHit "メール", place, MaskEmail(email), email
    Next item
End Sub

Private Sub FindLabeledNumbers(ByVal text As String, ByVal place As String, ByVal pattern As String, _
    ByVal kind As String, ByVal minDigits As Long, ByVal maxDigits As Long)

    Dim matches As Object
    Dim item As Object
    Dim token As String
    Dim digits As String
    Set matches = MatchAll(pattern, text)
    For Each item In matches
        If mTruncated Then Exit Sub
        token = CStr(item.Value)
        If Not IsDigitBounded(text, CLng(item.FirstIndex) + 1, Len(token)) Then GoTo NextToken
        digits = DigitsOf(token)
        If Len(digits) >= minDigits And Len(digits) <= maxDigits Then
            AddHit kind, place, MaskDigits(token, 2, 2), token
        End If
NextToken:
    Next item
End Sub

Private Sub FindPostals(ByVal text As String, ByVal place As String)
    Dim matches As Object
    Dim item As Object
    Dim token As String
    Dim startAt As Long
    Set matches = MatchAll("(?:〒\s*)?\d{3}-\d{4}", text)
    For Each item In matches
        If mTruncated Then Exit Sub
        token = CStr(item.Value)
        startAt = CLng(item.FirstIndex) + 1
        If IsPostalBounded(text, startAt, Len(token)) Then
            If Len(DigitsOf(token)) = 7 Then AddHit "郵便番号", place, token, token
        End If
    Next item
End Sub

Private Sub FindLongNumbers(ByVal text As String, ByVal place As String)
    Dim matches As Object
    Dim item As Object
    Dim token As String
    Dim digits As String
    Dim startAt As Long
    Set matches = MatchAll("(?:^|[^\d])((?:\d[ -]?){12,19})(?!\d)", text)
    For Each item In matches
        If mTruncated Then Exit Sub
        If item.SubMatches.Count = 0 Then GoTo NextToken
        token = CStr(item.SubMatches(0))
        startAt = InStr(1, CStr(item.Value), Left$(token, 1), vbBinaryCompare)
        If startAt = 0 Then startAt = 1
        startAt = CLng(item.FirstIndex) + startAt
        If Not IsDigitBounded(text, startAt, Len(token)) Then GoTo NextToken
        digits = DigitsOf(token)
        If Len(digits) >= 13 And Len(digits) <= 19 Then
            If LuhnOk(digits) And Not SameDigit(digits) Then AddHit "カード番号", place, MaskDigits(token, 0, 4), token
        ElseIf Len(digits) = 12 Then
            If MyNumberOk(digits) Then AddHit "個人番号", place, MaskDigits(token, 0, 2), token
        End If
NextToken:
    Next item
End Sub

Private Sub AddHit(ByVal kind As String, ByVal place As String, ByVal shown As String, ByVal rawMatch As String)
    Dim key As String
    If mTruncated Then Exit Sub
    If Len(mValueEcho) > 0 Then
        If InStr(1, mValueEcho, rawMatch, vbTextCompare) > 0 Then Exit Sub
    End If
    If mHits.Count >= MaxHits Then
        mTruncated = True
        Exit Sub
    End If
    key = kind & Chr$(1) & mSheet & Chr$(1) & place & Chr$(1) & shown
    If mSeen.Exists(key) Then Exit Sub
    mSeen.Add key, True
    mHits.Add Array(kind, mSheet, place, shown)
End Sub

Private Function MatchAll(ByVal pattern As String, ByVal text As String) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = pattern
    Set MatchAll = re.Execute(text)
End Function

Private Function MayContain(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch = "@" Or ch = "＠" Or ch = "〒" Then
            MayContain = True
            Exit Function
        End If
        If (ch >= "0" And ch <= "9") Or (ch >= "０" And ch <= "９") Then
            MayContain = True
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeScan(ByVal text As String) As String
    text = StrConv(text, vbNarrow)
    text = Replace(text, "ー", "-")
    text = Replace(text, "－", "-")
    text = Replace(text, "−", "-")
    text = Replace(text, "‐", "-")
    NormalizeScan = text
End Function

Private Function IsPostalBounded(ByVal text As String, ByVal startAt As Long, ByVal length As Long) As Boolean
    Dim ch As String
    If startAt > 1 Then
        ch = Mid$(text, startAt - 1, 1)
        If ch >= "0" And ch <= "9" Then Exit Function
        If ch = "-" Then Exit Function
    End If
    If startAt + length <= Len(text) Then
        ch = Mid$(text, startAt + length, 1)
        If ch >= "0" And ch <= "9" Then Exit Function
        If ch = "-" Then Exit Function
    End If
    IsPostalBounded = True
End Function

Private Function IsDigitBounded(ByVal text As String, ByVal startAt As Long, ByVal length As Long) As Boolean
    Dim ch As String
    If startAt > 1 Then
        ch = Mid$(text, startAt - 1, 1)
        If ch >= "0" And ch <= "9" Then Exit Function
    End If
    If startAt + length <= Len(text) Then
        ch = Mid$(text, startAt + length, 1)
        If ch >= "0" And ch <= "9" Then Exit Function
    End If
    IsDigitBounded = True
End Function

Private Function SameDigit(ByVal digits As String) As Boolean
    Dim i As Long
    If Len(digits) = 0 Then Exit Function
    For i = 2 To Len(digits)
        If Mid$(digits, i, 1) <> Left$(digits, 1) Then Exit Function
    Next i
    SameDigit = True
End Function

Private Function DigitsOf(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOf = DigitsOf & ch
    Next i
End Function

Private Function MaskEmail(ByVal email As String) As String
    Dim atPos As Long
    atPos = InStr(email, "@")
    If atPos <= 1 Then
        MaskEmail = "***"
    Else
        MaskEmail = Left$(email, 1) & "***" & Mid$(email, atPos)
    End If
End Function

Private Function MaskDigits(ByVal text As String, ByVal keepLeft As Long, ByVal keepRight As Long) As String
    Dim i As Long
    Dim ch As String
    Dim digitCount As Long
    Dim seen As Long
    Dim out As String
    digitCount = Len(DigitsOf(text))
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then
            seen = seen + 1
            If seen <= keepLeft Or seen > digitCount - keepRight Then
                out = out & ch
            Else
                out = out & "*"
            End If
        Else
            out = out & ch
        End If
    Next i
    MaskDigits = out
End Function

Private Function LuhnOk(ByVal digits As String) As Boolean
    Dim i As Long
    Dim total As Long
    Dim n As Long
    Dim alternate As Boolean
    If Len(digits) < 13 Then Exit Function
    alternate = False
    For i = Len(digits) To 1 Step -1
        n = Asc(Mid$(digits, i, 1)) - 48
        If alternate Then
            n = n * 2
            If n > 9 Then n = n - 9
        End If
        total = total + n
        alternate = Not alternate
    Next i
    LuhnOk = (total Mod 10 = 0)
End Function

Private Function MyNumberOk(ByVal digits As String) As Boolean
    Dim weights As Variant
    Dim i As Long
    Dim total As Long
    Dim remainder As Long
    Dim checkDigit As Long
    If Len(digits) <> 12 Then Exit Function
    weights = Array(6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    For i = 0 To 10
        total = total + (Asc(Mid$(digits, i + 1, 1)) - 48) * CLng(weights(i))
    Next i
    remainder = total Mod 11
    If remainder <= 1 Then
        checkDigit = 0
    Else
        checkDigit = 11 - remainder
    End If
    MyNumberOk = (checkDigit = Asc(Mid$(digits, 12, 1)) - 48)
End Function

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
    Dim shapeName As String
    If depth > 15 Or mTruncated Then Exit Sub
    shapeName = "図形"
    On Error Resume Next
    shapeName = shp.Name
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
    If Not chartObject Is Nothing Then ScanChart chartObject, shapeName
    Consider "図形 " & shapeName, caption
    Consider "図形 " & shapeName & " 代替テキスト", altText
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
    On Error GoTo 0
End Function

Private Sub ScanChart(ByVal chartObject As Chart, ByVal chartName As String)
    Dim axisItem As Axis
    Dim seriesItem As Series
    Dim text As String
    Dim pointIndex As Long
    Dim pointCount As Long
    On Error Resume Next
    text = ""
    If chartObject.HasTitle Then text = CStr(chartObject.ChartTitle.Text)
    If Len(text) > 0 Then Consider "グラフ " & chartName & " タイトル", text
    For Each axisItem In chartObject.Axes
        text = ""
        If axisItem.HasTitle Then text = CStr(axisItem.AxisTitle.Text)
        If Len(text) > 0 Then Consider "グラフ " & chartName & " 軸", text
    Next axisItem
    For Each seriesItem In chartObject.SeriesCollection
        text = CStr(seriesItem.Name)
        If Len(text) > 0 Then Consider "グラフ " & chartName & " 系列", text
        pointCount = 0
        If seriesItem.HasDataLabels Then pointCount = seriesItem.Points.Count
        If pointCount > 300 Then pointCount = 300
        For pointIndex = 1 To pointCount
            text = ""
            If seriesItem.Points(pointIndex).HasDataLabel Then
                text = CStr(seriesItem.Points(pointIndex).DataLabel.Text)
            End If
            If Len(text) > 0 Then Consider "グラフ " & chartName & " データラベル", text
        Next pointIndex
    Next seriesItem
    On Error GoTo 0
End Sub

Private Sub ScanNotes(ByVal ws As Worksheet)
    Dim noteItem As Comment
    For Each noteItem In ws.Comments
        If mTruncated Then Exit Sub
        Consider "メモ " & noteItem.Parent.Address(False, False), noteItem.Text
    Next noteItem
End Sub

Private Sub ScanThreadedComments(ByVal ws As Worksheet)
    Dim threads As Object
    Dim threadItem As Object
    Dim reply As Object
    Dim replies As Object
    Dim text As String
    Dim place As String
    On Error Resume Next
    Set threads = CallByName(ws, "CommentsThreaded", VbGet)
    On Error GoTo 0
    If threads Is Nothing Then Exit Sub
    For Each threadItem In threads
        If mTruncated Then Exit Sub
        text = ""
        place = ""
        On Error Resume Next
        text = CStr(threadItem.Text)
        place = threadItem.Parent.Address(False, False)
        Set replies = threadItem.Replies
        On Error GoTo 0
        Consider "コメント " & place, text
        If Not replies Is Nothing Then
            For Each reply In replies
                text = ""
                On Error Resume Next
                text = CStr(reply.Text)
                On Error GoTo 0
                Consider "コメント " & place, text
            Next reply
        End If
    Next threadItem
End Sub

Private Sub ScanLinkCollection(ByVal links As Hyperlinks, ByVal sheetName As String)
    Dim link As Hyperlink
    Dim place As String
    Dim savedSheet As String
    savedSheet = mSheet
    mSheet = sheetName
    For Each link In links
        If mTruncated Then Exit For
        place = ""
        On Error Resume Next
        If link.Type = msoHyperlinkRange Then
            place = link.Range.Address(False, False)
        Else
            place = link.Shape.Name
        End If
        On Error GoTo 0
        If Len(place) = 0 Then place = "リンク"
        Consider "ハイパーリンク " & place, link.Address
        Consider "ハイパーリンク " & place, link.TextToDisplay
    Next link
    mSheet = savedSheet
End Sub

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
        data(total, 4) = "該当は先頭 " & CStr(MaxHits) & " 件までです。"
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

Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function
