Attribute VB_Name = "modProof"
Option Explicit

' Excel の文字列を、インストール済みの Word の文章校正で調べる。
' 外部 API は使わない。参照設定は不要（Word は CreateObject で起動）。
' 日本語の校正ツールが入ったデスクトップ版 Word が必要。
' Word が見つける誤字と、ら抜き言葉などの文法指摘が対象。
' 対話式の「校閲 → 文章校正」より指摘は少ないことがある。
' このモジュールだけでインポートして使えます。

Private Const MaxHits As Long = 3000
Private Const MaxSuggestionNames As Long = 5
Private Const JapaneseLanguageId As Long = 1041
Private Const ReportSheetName As String = "文章確認"

Private mHits As Collection
Private mTruncated As Boolean
Private mWordApp As Object
Private mWordDoc As Object

' 【WordProofAvailable】デスクトップ版 Word を起動できるなら True。
' 使用例:
'   If Not WordProofAvailable() Then MsgBox "Word がありません。"
' 解説: Word が無い環境では校正できないので、先に確認する。起動できたあとはすぐ閉じる。
Public Function WordProofAvailable() As Boolean
    Dim app As Object
    On Error Resume Next
    Set app = CreateObject("Word.Application")
    If Not app Is Nothing Then
        WordProofAvailable = True
        app.Quit
        Set app = Nothing
    End If
    On Error GoTo 0
End Function

' 【CheckJapaneseText】1 件の文字列を日本語として校正し、指摘を二次元配列で返す。
'   列は 種別, 指摘, 候補。種別は 誤字 または 文法。問題が無ければ Empty。
'   候補は、Word が出せるときだけ。文法指摘では空になることが多い。
' 使用例:
'   hits = CheckJapaneseText("食べれる人が多いです。")
' 解説: 「食べれる」のようなら抜き言葉を文法として返す。正しい文なら Empty。
Public Function CheckJapaneseText(ByVal text As String) As Variant
    Dim openedHere As Boolean
    On Error GoTo EH
    If Len(Trim$(text)) = 0 Then Exit Function
    openedHere = EnsureWordSession()
    CheckJapaneseText = ProofOneText(text)
    If openedHere Then CloseWordSession
    Exit Function
EH:
    CloseWordSession
    Err.Raise Err.Number, "CheckJapaneseText", Err.Description
End Function

' 【FindJapaneseProofIssues】セル範囲の文字列を校正し、該当箇所を二次元配列で返す。
'   列は 種別, シート, 場所, 指摘, 候補。見つからなければ Empty。
'   文字列定数だけを見る。数式、数値、空欄は対象外。3000 件を超える分は省略する。
' 使用例:
'   hits = FindJapaneseProofIssues(Selection)
' 解説: 選んだ範囲の文字列セルを順に調べ、誤字と文法の一覧を返す。
Public Function FindJapaneseProofIssues(ByVal target As Range) As Variant
    Dim cell As Range
    Dim area As Range
    Dim textCells As Range
    Dim openedHere As Boolean
    On Error GoTo EH
    If target Is Nothing Then Err.Raise 5, "FindJapaneseProofIssues", "調べる範囲がありません。"
    Set textCells = TextConstants(target)
    If textCells Is Nothing Then Exit Function
    BeginHits
    openedHere = EnsureWordSession()
    For Each area In textCells.Areas
        For Each cell In area.Cells
            CollectCellIssues cell
            If mTruncated Then Exit For
        Next cell
        If mTruncated Then Exit For
    Next area
    FindJapaneseProofIssues = HitsToArray()
    ResetHits
    If openedHere Then CloseWordSession
    Exit Function
EH:
    ResetHits
    CloseWordSession
    Err.Raise Err.Number, "FindJapaneseProofIssues", Err.Description
End Function

' 【WriteJapaneseProofIssues】校正結果を destination へ書き、見出しを含む行数を返す。
'   該当が無ければ 2 行目に「（なし）」。= で始まる内容は文字として書く。
' 使用例:
'   rows = WriteJapaneseProofIssues(Worksheets("結果").Range("A1"), Selection)
' 解説: 選択範囲を調べ、「結果」の A1 から種別・シート・場所・指摘・候補を書く。
Public Function WriteJapaneseProofIssues(ByVal destination As Range, ByVal target As Range) As Long
    Dim hits As Variant
    Dim labels As Variant
    Dim i As Long
    Dim rows As Long
    If destination Is Nothing Then Err.Raise 5, "WriteJapaneseProofIssues", "書き込み先が指定されていません。"
    hits = FindJapaneseProofIssues(target)
    labels = Array("種別", "シート", "場所", "指摘", "候補")
    For i = 0 To 4
        destination.Cells(1, i + 1).Value = labels(i)
    Next i
    destination.Resize(1, 5).Font.Bold = True
    If IsEmpty(hits) Then
        destination.Cells(2, 1).Value = "（なし）"
        WriteJapaneseProofIssues = 2
        Exit Function
    End If
    rows = UBound(hits, 1)
    WriteHitRows destination, hits, rows
    WriteJapaneseProofIssues = rows + 1
End Function

' 【Run_CheckJapaneseSelection】マクロ一覧用。選択範囲の文字列を校正する。
'   結果は「文章確認」シートへ書く。実行前に、そのシートのセルは空にする。
' 使用例:
'   文言が入ったセル範囲を選択して実行する。
' 解説: 選んだ文字列を Word の日本語校正で調べ、誤字と文法の件数を表示する。Word が無いとメッセージを出して終わる。
Public Sub Run_CheckJapaneseSelection()
    Dim target As Range
    Dim wb As Workbook
    Dim report As Worksheet
    Dim written As Long
    Dim found As Long
    Dim prevScreen As Boolean
    Dim screenChanged As Boolean
    Set target = RequireSelection()
    If target Is Nothing Then Exit Sub
    On Error GoTo EH
    If Not WordProofAvailable() Then
        MsgBox "デスクトップ版 Word が見つかりません。日本語の校正ツールが入った Word をインストールしてから実行してください。", _
            vbExclamation, ReportSheetName
        Exit Sub
    End If
    Set wb = target.Worksheet.Parent
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenChanged = True
    Set report = EnsureSheet(wb, ReportSheetName)
    report.Cells.Clear
    written = WriteJapaneseProofIssues(report.Range("A1"), target)
    report.Columns("A:E").AutoFit
    report.Activate
    report.Range("A1").Select
    Application.ScreenUpdating = prevScreen
    If report.Range("A2").Value = "（なし）" Then
        MsgBox "誤字・文法の指摘はありませんでした。", vbInformation, ReportSheetName
    Else
        found = written - 1
        If report.Cells(written, 1).Value = "（省略）" Then found = found - 1
        MsgBox CStr(found) & " 件見つかりました。結果は「" & ReportSheetName & "」シートにあります。", _
            vbExclamation, ReportSheetName
    End If
    Exit Sub
EH:
    If screenChanged Then Application.ScreenUpdating = prevScreen
    MsgBox Err.Description, vbExclamation, ReportSheetName
End Sub

Private Function EnsureWordSession() As Boolean
    Dim created As Boolean
    On Error GoTo Fail
    If mWordApp Is Nothing Then
        Set mWordApp = CreateObject("Word.Application")
        created = True
        mWordApp.Visible = False
        mWordApp.DisplayAlerts = 0
        On Error Resume Next
        mWordApp.Options.CheckGrammarAsYouType = True
        mWordApp.Options.CheckSpellingAsYouType = True
        mWordApp.Options.CheckGrammarWithSpelling = True
        On Error GoTo Fail
        Set mWordDoc = mWordApp.Documents.Add
    End If
    EnsureWordSession = created
    Exit Function
Fail:
    CloseWordSession
    Err.Raise 5, "modProof", "Word を起動できません。デスクトップ版 Word と日本語の校正ツールを確認してください。"
End Function

Private Sub CloseWordSession()
    On Error Resume Next
    If Not mWordDoc Is Nothing Then
        mWordDoc.Close SaveChanges:=False
        Set mWordDoc = Nothing
    End If
    If Not mWordApp Is Nothing Then
        mWordApp.Quit
        Set mWordApp = Nothing
    End If
    On Error GoTo 0
End Sub

Private Function ProofOneText(ByVal text As String) As Variant
    Dim spell As Object
    Dim gram As Object
    Dim item As Object
    Dim bag As Collection
    Dim data() As Variant
    Dim i As Long
    Dim total As Long
    Set bag = New Collection
    ApplyText text
    Set spell = mWordDoc.SpellingErrors
    For Each item In spell
        bag.Add Array("誤字", ClipText(item.Text), SuggestionText(item))
    Next item
    Set gram = mWordDoc.GrammaticalErrors
    For Each item In gram
        bag.Add Array("文法", ClipText(item.Text), "")
    Next item
    If bag.Count = 0 Then Exit Function
    total = bag.Count
    ReDim data(1 To total, 1 To 3)
    For i = 1 To total
        data(i, 1) = bag(i)(0)
        data(i, 2) = bag(i)(1)
        data(i, 3) = bag(i)(2)
    Next i
    ProofOneText = data
End Function

Private Sub CollectCellIssues(ByVal cell As Range)
    Dim text As String
    Dim spell As Object
    Dim gram As Object
    Dim item As Object
    Dim place As String
    text = CStr(cell.Value)
    If Len(Trim$(text)) = 0 Then Exit Sub
    place = cell.Address(False, False)
    ApplyText text
    Set spell = mWordDoc.SpellingErrors
    For Each item In spell
        AddHit "誤字", cell.Worksheet.Name, place, ClipText(item.Text), SuggestionText(item)
        If mTruncated Then Exit Sub
    Next item
    Set gram = mWordDoc.GrammaticalErrors
    For Each item In gram
        AddHit "文法", cell.Worksheet.Name, place, ClipText(item.Text), ""
        If mTruncated Then Exit Sub
    Next item
End Sub

Private Sub ApplyText(ByVal text As String)
    Dim rng As Object
    mWordDoc.Content.Text = text
    Set rng = mWordDoc.Content
    rng.LanguageID = JapaneseLanguageId
    rng.NoProofing = False
End Sub

Private Function SuggestionText(ByVal errRange As Object) As String
    Dim list As Object
    Dim item As Object
    Dim names As String
    Dim count As Long
    On Error Resume Next
    Set list = errRange.GetSpellingSuggestions()
    If list Is Nothing Then
        Set list = mWordApp.GetSpellingSuggestions(errRange.Text)
    End If
    If list Is Nothing Then Exit Function
    For Each item In list
        count = count + 1
        If count > MaxSuggestionNames Then Exit For
        If Len(names) > 0 Then names = names & " / "
        names = names & CStr(item.Name)
    Next item
    On Error GoTo 0
    SuggestionText = names
End Function

Private Sub BeginHits()
    Set mHits = New Collection
    mTruncated = False
End Sub

Private Sub ResetHits()
    Set mHits = Nothing
    mTruncated = False
End Sub

Private Sub AddHit(ByVal kind As String, ByVal sheetName As String, ByVal place As String, _
    ByVal issue As String, ByVal suggestion As String)

    If mTruncated Then Exit Sub
    If mHits.Count >= MaxHits Then
        mTruncated = True
        Exit Sub
    End If
    mHits.Add Array(kind, sheetName, place, issue, suggestion)
End Sub

Private Function HitsToArray() As Variant
    Dim data() As Variant
    Dim index As Long
    Dim item As Variant
    Dim total As Long
    If mHits Is Nothing Or mHits.Count = 0 Then Exit Function
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
        data(total, 4) = "該当は先頭 " & CStr(MaxHits) & " 件までです。"
    End If
    HitsToArray = data
End Function

Private Sub WriteHitRows(ByVal destination As Range, ByVal hits As Variant, ByVal rows As Long)
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim text As String
    For rowIndex = 1 To rows
        For colIndex = 1 To 5
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

Private Function TextConstants(ByVal target As Range) As Range
    On Error Resume Next
    Set TextConstants = target.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
End Function

Private Function RequireSelection() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation, ReportSheetName
        Exit Function
    End If
    Set RequireSelection = Selection
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

Private Function ClipText(ByVal text As String) As String
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, Chr$(7), "")
    If Len(text) > 200 Then text = Left$(text, 200) & "…"
    ClipText = text
End Function
