Attribute VB_Name = "modSheet"
Option Explicit

' シートの有無確認、作成、削除、並び替え、保護。
' このモジュールだけでインポートして使えます。

' 【SheetExists】その名前のワークシートがあるか。チャートシートは見ない。
' 使用例:
'   If SheetExists("集計") Then Worksheets("集計").Activate
' 解説: 「集計」というワークシートがあるときだけ、そのシートを表示する。無いときに Worksheets("集計") へ触るとエラーになるので、先に確認している。
Public Function SheetExists(ByVal sheetName As String, Optional ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' 【GetSheet】シートを返す。無いときは Nothing なので、削除せず存在確認に使える。
' 使用例:
'   Set ws = GetSheet("集計")
'   If Not ws Is Nothing Then ws.Range("A1").Value = "更新"
' 解説: 「集計」があれば ws に取り、A1 へ「更新」と書く。シートが無いときは Nothing なので、何もせずに進む。
Public Function GetSheet(ByVal sheetName As String, Optional ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If SheetExists(sheetName, wb) Then Set GetSheet = wb.Worksheets(sheetName)
End Function

' 【GetOrCreateSheet】シートを返す。無ければブックの末尾に作る。
'   : \ / ? * [ ] はシート名に使えないので、先に SafeSheetName を通す。
' 使用例:
'   Set ws = GetOrCreateSheet("集計")
'   ws.Range("A1").Value = "日付"
' 解説: 「集計」があればそれを使い、無ければブックの末尾に新しく作ってから、A1 へ「日付」と書く。
Public Function GetOrCreateSheet(ByVal sheetName As String, Optional ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If SheetExists(sheetName, wb) Then
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
        Exit Function
    End If
    Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    GetOrCreateSheet.Name = sheetName
End Function

' 【DeleteSheet】シートを削除して True。元から無ければ False。
'   最後の 1 枚は削除できない。確認ダイアログは出さない。
' 使用例:
'   DeleteSheet "作業用"
' 解説: 前面のブックから「作業用」シートを、確認ダイアログなしで削除する。最初から無いときは何もしない。ブックにシートが 1 枚しか無いときはエラーになる。
Public Function DeleteSheet(ByVal sheetName As String, Optional ByVal wb As Workbook) As Boolean
    Dim prevAlerts As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not SheetExists(sheetName, wb) Then Exit Function
    If wb.Sheets.Count <= 1 Then Err.Raise 5, "DeleteSheet", "最後のシートは削除できません。"
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo EH
    wb.Worksheets(sheetName).Delete
    Application.DisplayAlerts = prevAlerts
    DeleteSheet = True
    Exit Function
EH:
    Application.DisplayAlerts = prevAlerts
    Err.Raise Err.Number, "DeleteSheet", Err.Description
End Function

' 【RenameSheet】シート名を変える。同名が既にあるとエラー。
' 使用例:
'   RenameSheet ThisWorkbook, "Sheet1", "売上"
' 解説: このマクロが入っているブックの Sheet1 を「売上」に変える。すでに「売上」があるときはエラーになる。
Public Sub RenameSheet(ByVal wb As Workbook, ByVal oldName As String, ByVal newName As String)
    If Not SheetExists(oldName, wb) Then Err.Raise 5, "RenameSheet", "シートがありません: " & oldName
    If StrComp(oldName, newName, vbTextCompare) <> 0 Then
        If SheetExists(newName, wb) Then Err.Raise 5, "RenameSheet", "同名のシートがあります: " & newName
    End If
    wb.Worksheets(oldName).Name = newName
End Sub

' 【SafeSheetName】シート名に使えない文字を _ にし、31 文字以内にする。
' 使用例:
'   ws.Name = SafeSheetName("2026/09 売上:速報")   ' "2026_09 売上_速報"
' 解説: スラッシュとコロンはシート名に使えないので _ に直し、「2026_09 売上_速報」という名前を付ける。31 文字を超えるときは後ろを切る。
Public Function SafeSheetName(ByVal sheetName As String) As String
    Dim ch As Variant
    For Each ch In Array(":", "\", "/", "?", "*", "[", "]")
        sheetName = Replace(sheetName, CStr(ch), "_")
    Next ch
    sheetName = Trim$(sheetName)
    If Len(sheetName) = 0 Then sheetName = "Sheet"
    If Len(sheetName) > 31 Then sheetName = Left$(sheetName, 31)
    SafeSheetName = sheetName
End Function

' 【UniqueSheetName】既存と重ならないシート名を返す。重なるときは _2, _3 を付ける。
' 使用例:
'   newName = UniqueSheetName(ActiveWorkbook, "集計")
'   Worksheets.Add.Name = newName
' 解説: 「集計」が無ければその名前、既にあれば「集計_2」のように空いている名前を作り、新しいシートへ付ける。
Public Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As String
    Dim i As Long
    candidate = SafeSheetName(baseName)
    If Not SheetExists(candidate, wb) Then
        UniqueSheetName = candidate
        Exit Function
    End If
    For i = 2 To 999
        suffix = "_" & CStr(i)
        candidate = Left$(SafeSheetName(baseName), 31 - Len(suffix)) & suffix
        If Not SheetExists(candidate, wb) Then
            UniqueSheetName = candidate
            Exit Function
        End If
    Next i
    Err.Raise 5, "UniqueSheetName", "空きシート名を作れません: " & baseName
End Function

' 【SheetNames】ワークシート名の配列。チャートシートは含まない。
'   includeVeryHidden が False のときは、非常に非表示のシートを除く。
' 使用例:
'   names = SheetNames()
'   For i = 1 To UBound(names)
'       Debug.Print names(i)
'   Next i
' 解説: 前面のブックのワークシート名を配列で受け取り、1 件ずつイミディエイトへ出す。グラフシートは含まない。非常に非表示のシートも、この呼び方では含む。
Public Function SheetNames(Optional ByVal wb As Workbook, Optional ByVal includeVeryHidden As Boolean = True) As Variant
    Dim ws As Worksheet
    Dim names() As String
    Dim count As Long
    If wb Is Nothing Then Set wb = ActiveWorkbook
    ReDim names(1 To wb.Worksheets.Count)
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVeryHidden Or includeVeryHidden Then
            count = count + 1
            names(count) = ws.Name
        End If
    Next ws
    If count = 0 Then
        SheetNames = Array()
    Else
        ReDim Preserve names(1 To count)
        SheetNames = names
    End If
End Function

' 【SortSheetsByName】ワークシートを名前の昇順に並べる。チャートシートは動かさない。
' 使用例:
'   SortSheetsByName ActiveWorkbook
' 解説: 前面のブックのワークシートを、名前の昇順に左から並べ替える。グラフシートの位置は変えない。
Public Sub SortSheetsByName(ByVal wb As Workbook)
    Dim names() As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    ReDim names(1 To wb.Worksheets.Count)
    For i = 1 To wb.Worksheets.Count
        names(i) = wb.Worksheets(i).Name
    Next i
    For i = 1 To UBound(names) - 1
        For j = i + 1 To UBound(names)
            If StrComp(names(i), names(j), vbTextCompare) > 0 Then
                tmp = names(i)
                names(i) = names(j)
                names(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To UBound(names)
        wb.Worksheets(names(i)).Move Before:=wb.Worksheets(i)
    Next i
End Sub

' 【ProtectAllSheets】全ワークシートを保護する。
'   UserInterfaceOnly により、開いている間はマクロから編集できる。
'   開き直すとこの指定は消えるので、必要なときは Workbook_Open でかけ直す。
' 使用例:
'   ProtectAllSheets ThisWorkbook, "secret"
'   ProtectAllSheets ActiveWorkbook          ' パスワードなし
' 解説: 1 行目はこのマクロのブックをパスワード secret で保護する。2 行目は前面のブックを、パスワードなしで保護する。開いている間はマクロからセルを編集できる。
Public Sub ProtectAllSheets(ByVal wb As Workbook, Optional ByVal password As String = "", _
    Optional ByVal allowFormattingCells As Boolean = True)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.Protect Password:=password, UserInterfaceOnly:=True, DrawingObjects:=False, _
            Contents:=True, AllowFormattingCells:=allowFormattingCells
    Next ws
End Sub

' 【UnprotectAllSheets】全ワークシートの保護を解除する。パスワードが違うとエラー。
' 使用例:
'   UnprotectAllSheets ThisWorkbook, "secret"
' 解説: このマクロのブックの全シートから、パスワード secret の保護を外す。違うパスワードのシートがあると、そこでエラーになる。
Public Sub UnprotectAllSheets(ByVal wb As Workbook, Optional ByVal password As String = "")
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=password
        If Err.Number <> 0 Then
            Err.Raise 5, "UnprotectAllSheets", "保護を解除できませんでした: " & ws.Name
        End If
        On Error GoTo 0
    Next ws
End Sub

' 【ResetViewToA1】表示中のシートを A1 に戻し、スクロールも左上にする。
' 使用例:
'   ResetViewToA1 ThisWorkbook
' 解説: このブックの表示中シートを、選択セルもスクロールも A1 に戻す。非表示シートは触らない。終わると、もともと開いていたシートに戻る。
Public Sub ResetViewToA1(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim returnTo As Worksheet
    Dim prevScreen As Boolean
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo EH
    Set returnTo = ActiveSheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            Application.Goto ws.Range("A1"), True
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next ws
    If Not returnTo Is Nothing Then
        If returnTo.Visible = xlSheetVisible Then returnTo.Activate
    End If
    Application.ScreenUpdating = prevScreen
    Exit Sub
EH:
    Application.ScreenUpdating = prevScreen
    Err.Raise Err.Number, "ResetViewToA1", Err.Description
End Sub

' 【ShowAllSheets】非表示と「非常に非表示」を含め、すべてのシートを表示する。
' 使用例:
'   ShowAllSheets ActiveWorkbook
' 解説: 前面のブックで、隠したシートと「非常に非表示」のシートをすべて表示する。
Public Sub ShowAllSheets(ByVal wb As Workbook)
    Dim sh As Object
    For Each sh In wb.Sheets
        sh.Visible = xlSheetVisible
    Next sh
End Sub

' 【SetVeryHidden】シートを「非常に非表示」にする。ユーザーは右クリックから再表示できない。
'   表示を戻すときは ShowAllSheets か、ws.Visible = xlSheetVisible。
' 使用例:
'   SetVeryHidden Worksheets("設定")
' 解説: 「設定」シートを非常に非表示にする。シート見出しの右クリックからは再表示できない。戻すときは ShowAllSheets を使う。
Public Sub SetVeryHidden(ByVal ws As Worksheet)
    ws.Visible = xlSheetVeryHidden
End Sub

' 【CopySheetToNewWorkbook】シートを新しいブックへ複製し、そのブックを返す。
' 使用例:
'   Set copied = CopySheetToNewWorkbook(Worksheets("提出用"))
'   copied.SaveAs "C:\temp\提出用.xlsx"
' 解説: 「提出用」シートだけを新しいブックへ複製し、C:\temp\提出用.xlsx として保存する。元のブックのシートはそのまま残る。
Public Function CopySheetToNewWorkbook(ByVal ws As Worksheet) As Workbook
    ws.Copy
    Set CopySheetToNewWorkbook = ActiveWorkbook
End Function

' 【Run_SortSheetNames】マクロ一覧用。開いているブックのシートを名前順に並べる。
' 使用例:
'   並べ替えたいブックを前面にして実行する。
' 解説: 前面のブックのシート見出しが、名前の昇順に並び替わる。グラフシートは動かない。
Public Sub Run_SortSheetNames()
    On Error GoTo EH
    SortSheetsByName ActiveWorkbook
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "SortSheetsByName"
End Sub

' 【Run_ResetViewToA1】マクロ一覧用。表示中の全シートを A1 表示に戻す。
' 使用例:
'   納品前に、スクロール位置を揃えるために実行する。
' 解説: 表示中の全シートが A1 を左上にした状態に戻る。受け取った人が途中の行から開かないようにする。
Public Sub Run_ResetViewToA1()
    On Error GoTo EH
    ResetViewToA1 ActiveWorkbook
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "ResetViewToA1"
End Sub

' 【Run_ShowAllSheets】マクロ一覧用。隠したシートをすべて表示する。
' 使用例:
'   非表示シートをまとめて戻したいときに実行する。
' 解説: 前面のブックで隠れているシートが、すべてシート見出しに戻る。非常に非表示のものも含む。
Public Sub Run_ShowAllSheets()
    On Error GoTo EH
    ShowAllSheets ActiveWorkbook
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "ShowAllSheets"
End Sub

' 【Run_ProtectAllSheets】マクロ一覧用。パスワードを聞いて全シートを保護する。
'   キャンセルすると何もしない。空のまま OK するとパスワードなしで保護する。
' 使用例:
'   配布前に実行し、編集されたくないシートをまとめて保護する。
' 解説: 入力欄にパスワードを入れて OK すると、全シートが保護される。空のまま OK するとパスワードなし。キャンセルすると何もしない。
Public Sub Run_ProtectAllSheets()
    Dim password As String
    password = InputBox("シート保護のパスワード（空のままでも保護できます）", "シート保護")
    If StrPtr(password) = 0 Then Exit Sub
    On Error GoTo EH
    ProtectAllSheets ActiveWorkbook, password
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "ProtectAllSheets"
End Sub

' 【Run_UnprotectAllSheets】マクロ一覧用。パスワードを聞いて全シートの保護を解除する。
'   パスワードが違うシートがあると、そこでメッセージを出して止まる。
' 使用例:
'   保護をまとめて外してからデータを直す。設定が無ければ空のまま OK。
' 解説: 保護したときと同じパスワードを入れると、全シートの保護が外れる。違うパスワードのシートがあると、そこでメッセージが出て止まる。
Public Sub Run_UnprotectAllSheets()
    Dim password As String
    password = InputBox("シート保護のパスワード（設定していなければ空のまま OK）", "保護解除")
    If StrPtr(password) = 0 Then Exit Sub
    On Error GoTo EH
    UnprotectAllSheets ActiveWorkbook, password
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "UnprotectAllSheets"
End Sub
