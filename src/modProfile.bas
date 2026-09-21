Attribute VB_Name = "modProfile"
Option Explicit

' ドキュメントプロパティ、ヘッダーとフッター、外部リンク、接続、クエリを一覧する。
' このモジュールだけでインポートして使えます。
' ヘッダーの読み取りはプリンタと通信しないよう、印刷通信を止めてから行う。
' 接続文字列の Password や Pwd は *** に置き換える。

Private mHits As Collection
Private mPrintChanged As Boolean
Private mPrevPrint As Boolean

' 【FindWorkbookProfile】ブックの付帯情報を二次元配列で返す。
'   列は 区分, シート, 場所, 内容。見つからなければ Empty。
'   区分は ブック, プロパティ, ユーザー定義, ヘッダー, フッター, 外部リンク, 接続, クエリ, ピボット。
' 使用例:
'   hits = FindWorkbookProfile(ActiveWorkbook)
' 解説: 前面のブックの作成者や会社名、各シートのヘッダーとフッター、他ブックへのリンク、Power Query などの接続を集める。パスワードらしい値は伏せる。
Public Function FindWorkbookProfile(Optional ByVal wb As Workbook) As Variant
    On Error GoTo EH
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise 5, "FindWorkbookProfile", "調べるブックがありません。"
    Set mHits = New Collection
    QuietPrinting
    AddHit "ブック", "", "ファイル名", wb.Name
    If Len(wb.Path) > 0 Then AddHit "ブック", "", "パス", wb.FullName
    ScanBuiltinProperties wb
    ScanCustomProperties wb
    ScanSetups wb
    ScanExternalLinks wb
    ScanConnections wb
    ScanQueries wb
    ScanPivotCaches wb
    FindWorkbookProfile = HitsToArray()
    RestorePrinting
    Set mHits = Nothing
    Exit Function
EH:
    RestorePrinting
    Set mHits = Nothing
    Err.Raise Err.Number, "FindWorkbookProfile", Err.Description
End Function

' 【WriteWorkbookProfile】付帯情報を destination へ書き、見出しを含む行数を返す。
'   該当が無ければ 2 行目に「（なし）」。
' 使用例:
'   rows = WriteWorkbookProfile(Worksheets("結果").Range("A1"), ActiveWorkbook)
' 解説: 前面のブックの付帯情報を、「結果」の A1 から区分・シート・場所・内容で書く。
Public Function WriteWorkbookProfile(ByVal destination As Range, Optional ByVal wb As Workbook) As Long
    Dim hits As Variant
    Dim labels As Variant
    Dim i As Long
    Dim rows As Long
    If destination Is Nothing Then Err.Raise 5, "WriteWorkbookProfile", "書き込み先が指定されていません。"
    hits = FindWorkbookProfile(wb)
    labels = Array("区分", "シート", "場所", "内容")
    For i = 0 To 3
        destination.Cells(1, i + 1).Value = labels(i)
    Next i
    destination.Resize(1, 4).Font.Bold = True
    If IsEmpty(hits) Then
        destination.Cells(2, 1).Value = "（なし）"
        WriteWorkbookProfile = 2
        Exit Function
    End If
    rows = UBound(hits, 1)
    WriteHitRows destination, hits, rows
    WriteWorkbookProfile = rows + 1
End Function

' 【Run_WriteWorkbookProfile】マクロ一覧用。前面のブックの付帯情報を「付帯情報」シートへ書く。
'   実行前に、そのシートのセルは空にする。ヘッダーの読み取りには少し時間がかかることがある。
' 使用例:
'   提出前のブックを前面にして実行する。
' 解説: 作成者、ヘッダーとフッター、外部リンク、接続、クエリを1枚にまとめ、件数を表示する。
Public Sub Run_WriteWorkbookProfile()
    Dim wb As Workbook
    Dim report As Worksheet
    Dim written As Long
    Dim found As Long
    Dim prevScreen As Boolean
    Dim screenChanged As Boolean
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "調べるブックがありません。", vbExclamation, "付帯情報"
        Exit Sub
    End If
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenChanged = True
    Set report = EnsureSheet(wb, "付帯情報")
    report.Cells.Clear
    written = WriteWorkbookProfile(report.Range("A1"), wb)
    report.Columns("A:D").AutoFit
    report.Activate
    report.Range("A1").Select
    Application.ScreenUpdating = prevScreen
    If report.Range("A2").Value = "（なし）" Then
        MsgBox "付帯情報は見つかりませんでした。", vbInformation, "付帯情報"
    Else
        found = written - 1
        MsgBox CStr(found) & " 件を「付帯情報」シートに書きました。", vbInformation, "付帯情報"
    End If
    Exit Sub
EH:
    If screenChanged Then Application.ScreenUpdating = prevScreen
    MsgBox Err.Description, vbExclamation, "付帯情報"
End Sub

Private Sub ScanBuiltinProperties(ByVal wb As Workbook)
    Dim keys As Variant
    Dim labels As Variant
    Dim i As Long
    Dim text As String
    keys = Array("Title", "Subject", "Author", "Last author", "Company", "Manager", "Comments", "Category", "Keywords", "Hyperlink base")
    labels = Array("タイトル", "件名", "作成者", "前回保存者", "会社", "管理者", "コメント", "分類", "キーワード", "ハイパーリンクの基点")
    For i = 0 To UBound(keys)
        text = ""
        On Error Resume Next
        text = Trim$(CStr(wb.BuiltinDocumentProperties(CStr(keys(i))).Value))
        On Error GoTo 0
        If Len(text) > 0 Then AddHit "プロパティ", "", CStr(labels(i)), RedactSecrets(text)
    Next i
End Sub

Private Sub ScanCustomProperties(ByVal wb As Workbook)
    Dim prop As DocumentProperty
    Dim text As String
    On Error Resume Next
    For Each prop In wb.CustomDocumentProperties
        text = ""
        text = Trim$(CStr(prop.Value))
        If Len(text) > 0 Then AddHit "ユーザー定義", "", prop.Name, RedactSecrets(text)
    Next prop
    On Error GoTo 0
End Sub

Private Sub ScanSetups(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim chartSheet As Chart
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then ReadSetup ws, ws.Name
    Next ws
    For Each chartSheet In wb.Charts
        If chartSheet.Visible = xlSheetVisible Then ReadSetup chartSheet, chartSheet.Name
    Next chartSheet
End Sub

Private Sub ReadSetup(ByVal sheetObject As Object, ByVal sheetName As String)
    ReadSetupMember sheetObject, sheetName, "ヘッダー", "左", "LeftHeader"
    ReadSetupMember sheetObject, sheetName, "ヘッダー", "中央", "CenterHeader"
    ReadSetupMember sheetObject, sheetName, "ヘッダー", "右", "RightHeader"
    ReadSetupMember sheetObject, sheetName, "フッター", "左", "LeftFooter"
    ReadSetupMember sheetObject, sheetName, "フッター", "中央", "CenterFooter"
    ReadSetupMember sheetObject, sheetName, "フッター", "右", "RightFooter"
    ReadSpecialSetup sheetObject, sheetName, "FirstPage", "先頭ページ"
    ReadSpecialSetup sheetObject, sheetName, "EvenPage", "偶数ページ"
End Sub

Private Sub ReadSpecialSetup(ByVal sheetObject As Object, ByVal sheetName As String, ByVal pageName As String, ByVal prefix As String)
    Dim pageObject As Object
    On Error Resume Next
    Set pageObject = CallByName(sheetObject.PageSetup, pageName, VbGet)
    On Error GoTo 0
    If pageObject Is Nothing Then Exit Sub
    ReadMember pageObject, sheetName, "ヘッダー", prefix & "左", "LeftHeader"
    ReadMember pageObject, sheetName, "ヘッダー", prefix & "中央", "CenterHeader"
    ReadMember pageObject, sheetName, "ヘッダー", prefix & "右", "RightHeader"
    ReadMember pageObject, sheetName, "フッター", prefix & "左", "LeftFooter"
    ReadMember pageObject, sheetName, "フッター", prefix & "中央", "CenterFooter"
    ReadMember pageObject, sheetName, "フッター", prefix & "右", "RightFooter"
End Sub

Private Sub ReadSetupMember(ByVal sheetObject As Object, ByVal sheetName As String, ByVal kind As String, ByVal place As String, ByVal memberName As String)
    On Error Resume Next
    ReadMember sheetObject.PageSetup, sheetName, kind, place, memberName
    On Error GoTo 0
End Sub

Private Sub ReadMember(ByVal target As Object, ByVal sheetName As String, ByVal kind As String, ByVal place As String, ByVal memberName As String)
    Dim text As String
    On Error Resume Next
    text = CStr(CallByName(target, memberName, VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    text = Trim$(Replace(Replace(text, vbCr, " "), vbLf, " "))
    If Len(text) > 0 Then AddHit kind, sheetName, place, RedactSecrets(Left$(text, 300))
End Sub

Private Sub ScanExternalLinks(ByVal wb As Workbook)
    Dim links As Variant
    Dim i As Long
    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0
    If IsEmpty(links) Then Exit Sub
    If Not IsArray(links) Then Exit Sub
    For i = LBound(links) To UBound(links)
        AddHit "外部リンク", "", "リンク", CStr(links(i))
    Next i
End Sub

Private Sub ScanConnections(ByVal wb As Workbook)
    Dim conn As WorkbookConnection
    Dim summary As String
    On Error Resume Next
    For Each conn In wb.Connections
        summary = ConnectionSummary(conn)
        AddHit "接続", "", conn.Name, RedactSecrets(summary)
    Next conn
    On Error GoTo 0
End Sub

Private Function ConnectionSummary(ByVal conn As WorkbookConnection) As String
    Dim parts As String
    On Error Resume Next
    parts = CStr(conn.OLEDBConnection.Connection)
    If Len(conn.OLEDBConnection.CommandText) > 0 Then parts = parts & " | " & CStr(conn.OLEDBConnection.CommandText)
    If Len(parts) = 0 Then parts = CStr(conn.ODBCConnection.Connection)
    If Len(parts) = 0 And Len(conn.ODBCConnection.CommandText) > 0 Then parts = CStr(conn.ODBCConnection.CommandText)
    If Len(parts) = 0 Then parts = CStr(conn.Description)
    On Error GoTo 0
    If Len(parts) = 0 Then parts = "接続の詳細は取得できませんでした。"
    ConnectionSummary = Left$(parts, 500)
End Function

Private Sub ScanQueries(ByVal wb As Workbook)
    Dim queries As Object
    Dim queryItem As Object
    Dim formulaText As String
    Dim queryName As String
    On Error Resume Next
    Set queries = CallByName(wb, "Queries", VbGet)
    On Error GoTo 0
    If queries Is Nothing Then Exit Sub
    For Each queryItem In queries
        formulaText = ""
        queryName = ""
        On Error Resume Next
        queryName = CStr(queryItem.Name)
        formulaText = CStr(queryItem.Formula)
        On Error GoTo 0
        AddHit "クエリ", "", queryName, RedactSecrets(Left$(formulaText, 500))
    Next queryItem
End Sub

Private Sub ScanPivotCaches(ByVal wb As Workbook)
    Dim cache As PivotCache
    Dim source As String
    Dim index As Long
    For Each cache In wb.PivotCaches
        index = index + 1
        source = ""
        On Error Resume Next
        source = CStr(cache.SourceData)
        If Len(source) = 0 Then
            If Not cache.WorkbookConnection Is Nothing Then source = cache.WorkbookConnection.Name
        End If
        On Error GoTo 0
        If IsExternalSource(source) Then
            AddHit "ピボット", "", "キャッシュ" & CStr(index), RedactSecrets(Left$(source, 500))
        End If
    Next cache
End Sub

Private Function IsExternalSource(ByVal source As String) As Boolean
    If Len(source) = 0 Then Exit Function
    IsExternalSource = (InStr(source, "[") > 0 Or InStr(source, ":\") > 0 Or InStr(source, "\\") > 0 Or InStr(1, source, "http", vbTextCompare) > 0)
End Function

Private Function RedactSecrets(ByVal text As String) As String
    Dim re As Object
    If Len(text) = 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(Password|Pwd|User ID|UID|AccountKey|SharedAccessSignature|sig)\s*=\s*[^;""&\s]*"
    RedactSecrets = re.Replace(text, "$1=***")
End Function

Private Sub QuietPrinting()
    On Error Resume Next
    mPrevPrint = Application.PrintCommunication
    Application.PrintCommunication = False
    mPrintChanged = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub RestorePrinting()
    If Not mPrintChanged Then Exit Sub
    On Error Resume Next
    Application.PrintCommunication = mPrevPrint
    Err.Clear
    On Error GoTo 0
    mPrintChanged = False
End Sub

Private Sub AddHit(ByVal kind As String, ByVal sheetName As String, ByVal place As String, ByVal content As String)
    If mHits Is Nothing Then Exit Sub
    If Len(content) = 0 And kind <> "接続" And kind <> "クエリ" Then Exit Sub
    mHits.Add Array(kind, sheetName, place, content)
End Sub

Private Function HitsToArray() As Variant
    Dim data() As Variant
    Dim index As Long
    Dim item As Variant
    If mHits Is Nothing Or mHits.Count = 0 Then Exit Function
    ReDim data(1 To mHits.Count, 1 To 4)
    For index = 1 To mHits.Count
        item = mHits(index)
        data(index, 1) = item(0)
        data(index, 2) = item(1)
        data(index, 3) = item(2)
        data(index, 4) = item(3)
    Next index
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

Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function
