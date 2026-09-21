Attribute VB_Name = "modBook"
Option Explicit

' ブックを開く、保存する、外部リンクとクエリを確認する。
' このモジュールだけでインポートして使えます。

' 【FindWorkbook】開いているブックを、フルパスまたはファイル名で探す。
'   パスを渡したときはパスが一致するものだけ。ファイル名だけなら同名の開いているブック。
'   見つからないときは Nothing。
' 使用例:
'   Set wb = FindWorkbook("C:\data\売上.xlsx")
'   Set wb = FindWorkbook("売上.xlsx")
' 解説: 1 行目は C:\data\売上.xlsx というパスで開いているブックだけを探す。2 行目はファイル名が売上.xlsx の開いているブックを探す。どちらも無ければ wb は Nothing。
Public Function FindWorkbook(ByVal pathOrName As String) As Workbook
    Dim wb As Workbook
    Dim leaf As String
    Dim matchPath As Boolean
    If Len(pathOrName) = 0 Then Exit Function
    matchPath = (InStr(pathOrName, "\") > 0)
    leaf = Mid$(pathOrName, InStrRev(pathOrName, "\") + 1)
    For Each wb In Application.Workbooks
        If matchPath Then
            If StrComp(wb.FullName, pathOrName, vbTextCompare) = 0 Then
                Set FindWorkbook = wb
                Exit Function
            End If
        ElseIf StrComp(wb.Name, leaf, vbTextCompare) = 0 Then
            Set FindWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

' 【IsWorkbookOpen】指定のブックが開いているか。
' 使用例:
'   If Not IsWorkbookOpen("マスタ.xlsx") Then
'       Set wb = OpenWorkbook("C:\data\マスタ.xlsx")
'   End If
' 解説: マスタ.xlsx がまだ開いていなければ、C:\data\マスタ.xlsx を開いて wb に入れる。開いていれば何もしない。
Public Function IsWorkbookOpen(ByVal pathOrName As String) As Boolean
    IsWorkbookOpen = Not FindWorkbook(pathOrName) Is Nothing
End Function

' 【OpenWorkbook】ブックを開いて返す。同じパスですでに開いていればそれを返す。
'   readOnly が True のときは読み取り専用。リンクの更新確認は出さない。
' 使用例:
'   Set wb = OpenWorkbook(CombinePath(ThisFolder(), "マスタ.xlsx"), True)
' 解説: このマクロのブックと同じフォルダのマスタ.xlsx を、読み取り専用で開く。すでに開いていれば、開き直さずにそのブックを返す。
Public Function OpenWorkbook(ByVal filePath As String, Optional ByVal readOnly As Boolean = False) As Workbook
    Dim wb As Workbook
    If Len(Dir$(filePath)) = 0 Then Err.Raise 53, "OpenWorkbook", "ファイルが見つかりません: " & filePath
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set OpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set OpenWorkbook = Workbooks.Open(Filename:=filePath, ReadOnly:=readOnly, UpdateLinks:=0)
End Function

' 【ThisFolder】このマクロが入っているブックのフォルダ。未保存なら空文字。
' 使用例:
'   outPath = CombinePath(ThisFolder(), "出力.csv")
' 解説: マクロが入っているブックのフォルダに「出力.csv」を足したフルパスを作る。ブックが未保存だとフォルダが空なので、パスは「出力.csv」だけになる。
Public Function ThisFolder() As String
    ThisFolder = ThisWorkbook.Path
End Function

' 【ActiveFolder】前面にあるブックのフォルダ。未保存なら空文字。
' 使用例:
'   If Len(ActiveFolder()) = 0 Then MsgBox "先にブックを保存してください。"
' 解説: 前面のブックが一度も保存されていないときは、保存を促すメッセージを出す。保存済みならフォルダパスが取れるので、メッセージは出ない。
Public Function ActiveFolder() As String
    ActiveFolder = ActiveWorkbook.Path
End Function

' 【SaveAllWorkbooks】保存済みで、読み取り専用でもアドインでもないブックを保存する。
'   戻り値は保存した件数。未保存の新規ブックは保存ダイアログを出さない。
' 使用例:
'   saved = SaveAllWorkbooks()
' 解説: 保存先が決まっているブックをまとめて上書き保存し、保存した件数を saved に入れる。新規未保存、読み取り専用、アドインは対象外。
Public Function SaveAllWorkbooks() As Long
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Len(wb.Path) > 0 And Not wb.ReadOnly And Not wb.IsAddin Then
            wb.Save
            SaveAllWorkbooks = SaveAllWorkbooks + 1
        End If
    Next wb
End Function

' 【CloseOtherWorkbooks】このマクロが入っているブック以外を閉じる。マクロ一覧には出ない。
'   saveChanges が True なら変更を保存、False なら破棄する。
' 使用例:
'   CloseOtherWorkbooks False
' 解説: このマクロのブック以外を、変更を保存せずに閉じる。保存して閉じるときは False の代わりに True を渡す。
Public Sub CloseOtherWorkbooks(ByVal saveChanges As Boolean)
    Dim wb As Workbook
    Dim keep As Workbook
    Set keep = ThisWorkbook
    Application.DisplayAlerts = False
    On Error GoTo EH
    For Each wb In Application.Workbooks
        If Not wb Is keep And Not wb.IsAddin Then
            wb.Close SaveChanges:=saveChanges
        End If
    Next wb
    Application.DisplayAlerts = True
    Exit Sub
EH:
    Application.DisplayAlerts = True
    Err.Raise Err.Number, "CloseOtherWorkbooks", Err.Description
End Sub

' 【RefreshWorkbook】Power Query などの更新と、ピボットテーブルの更新を行う。
'   非同期クエリは完了まで待つ。
' 使用例:
'   RefreshWorkbook ThisWorkbook
' 解説: このブックの Power Query などを更新し、終わるまで待ってからピボットテーブルも更新する。
Public Sub RefreshWorkbook(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim pivot As PivotTable
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo EH
    wb.RefreshAll
    For Each ws In wb.Worksheets
        For Each pivot In ws.PivotTables
            pivot.RefreshTable
        Next pivot
    Next ws
    Application.CalculateUntilAsyncQueriesDone
    Application.DisplayAlerts = prevAlerts
    Exit Sub
EH:
    Application.DisplayAlerts = prevAlerts
    Err.Raise Err.Number, "RefreshWorkbook", Err.Description
End Sub

' 【ExternalLinks】外部リンクの配列。リンクが無ければ要素 0 の配列。
' 使用例:
'   links = ExternalLinks(ThisWorkbook)
'   If UBound(links) >= LBound(links) Then Debug.Print links(LBound(links))
' 解説: このブックが参照している他ブックの一覧を取り、1 件でもあれば先頭のリンク先をイミディエイトへ出す。リンクが無ければ何も出さない。
Public Function ExternalLinks(ByVal wb As Workbook) As Variant
    Dim links As Variant
    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0
    If IsEmpty(links) Then
        ExternalLinks = Array()
    Else
        ExternalLinks = links
    End If
End Function

' 【Run_SaveAllWorkbooks】マクロ一覧用。保存済みの開いているブックをまとめて保存する。
'   未保存の新規ブックと、読み取り専用、アドインは対象外。終わると件数を表示する。
' 使用例:
'   作業中のブックをまとめて保存したいときに実行する。
' 解説: 保存済みのブックが上書き保存され、終わると「何件保存したか」のメッセージが出る。一度も保存していない新規ブックは対象外。
Public Sub Run_SaveAllWorkbooks()
    Dim saved As Long
    On Error GoTo EH
    saved = SaveAllWorkbooks()
    MsgBox CStr(saved) & " 件のブックを保存しました。", vbInformation, "SaveAllWorkbooks"
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "SaveAllWorkbooks"
End Sub

' 【Run_RefreshWorkbook】マクロ一覧用。前面のブックのクエリとピボットを更新する。
' 使用例:
'   データソースを更新したあとに実行し、表とピボットを最新にする。
' 解説: 前面のブックのクエリとピボットが更新され、終わると完了メッセージが出る。失敗したときはエラーの内容を表示する。
Public Sub Run_RefreshWorkbook()
    On Error GoTo EH
    RefreshWorkbook ActiveWorkbook
    MsgBox "更新が完了しました。", vbInformation, "RefreshWorkbook"
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "RefreshWorkbook"
End Sub
