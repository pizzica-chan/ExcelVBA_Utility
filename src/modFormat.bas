Attribute VB_Name = "modFormat"
Option Explicit

' 見出し、罫線、列幅、ウィンドウ枠の固定、印刷設定。
' このモジュールだけでインポートして使えます。
' 印刷設定はプリンタドライバと通信するため、初回は時間がかかることがあります。

' 【StyleHeaderRow】見出し行を白文字・紺背景・中央揃えにする。行の高さは 22。
'   fillColor を省略すると紺。fontColor を省略すると白。
' 使用例:
'   StyleHeaderRow Worksheets("一覧").Range("A1:F1")
'   StyleHeaderRow Range("A1:D1"), RGB(0, 112, 192), vbWhite
' 解説: 1 行目は「一覧」の A1:F1 を白文字・紺背景・中央揃えにする。2 行目は A1:D1 を、指定した青背景と白文字にする。行の高さはどちらも 22。
Public Sub StyleHeaderRow(ByVal header As Range, _
    Optional ByVal fillColor As Long = -1, _
    Optional ByVal fontColor As Long = 16777215)

    If fillColor < 0 Then fillColor = RGB(31, 78, 121)
    With header
        .Font.Bold = True
        .Font.Color = fontColor
        .Interior.Color = fillColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    header.RowHeight = 22
End Sub

' 【ApplyGridBorders】範囲に薄い罫線を引く。色を省略すると灰色。
' 使用例:
'   ApplyGridBorders Worksheets("一覧").Range("A1:F20")
' 解説: 「一覧」の A1:F20 に、灰色の細い罫線を引く。
Public Sub ApplyGridBorders(ByVal target As Range, Optional ByVal borderColor As Long = -1)
    If borderColor < 0 Then borderColor = RGB(191, 191, 191)
    target.Borders.LineStyle = xlContinuous
    target.Borders.Color = borderColor
    target.Borders.Weight = xlThin
End Sub

' 【StripeRows】偶数行に薄い色を付ける条件付き書式を追加する。
' 使用例:
'   StripeRows Worksheets("一覧").Range("A2:F200")
' 解説: 「一覧」の A2:F200 で、偶数行に薄い灰色の条件付き書式を足す。既存の条件付き書式は消さない。
Public Sub StripeRows(ByVal target As Range, Optional ByVal fillColor As Long = -1)
    If fillColor < 0 Then fillColor = RGB(242, 242, 242)
    With target.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0")
        .Interior.Color = fillColor
        .StopIfTrue = False
    End With
End Sub

' 【AutoFitLimited】列幅を内容に合わせ、狭すぎ・広すぎを minWidth と maxWidth に収める。
'   既定は幅 8 から 40。対象列のシート全体の内容で幅が決まる。
' 使用例:
'   AutoFitLimited Worksheets("一覧").Range("A:F"), 10, 30
' 解説: 「一覧」の A から F 列を内容に合わせ、列幅が 10 より狭いときは 10、30 より広いときは 30 に収める。
Public Sub AutoFitLimited(ByVal target As Range, Optional ByVal minWidth As Double = 8, Optional ByVal maxWidth As Double = 40)
    Dim column As Range
    target.Columns.AutoFit
    For Each column In target.Columns
        If column.ColumnWidth < minWidth Then column.ColumnWidth = minWidth
        If column.ColumnWidth > maxWidth Then column.ColumnWidth = maxWidth
    Next column
End Sub

' 【FreezePanesAt】指定セルの上と左でウィンドウ枠を固定する。
'   先頭行だけ固定するなら A2、先頭行と A 列なら B2 を渡す。
' 使用例:
'   FreezePanesAt Worksheets("一覧").Range("B2")
' 解説: 「一覧」で、1 行目と A 列を固定する。スクロールしても見出し行と左端の列が残る。
Public Sub FreezePanesAt(ByVal cell As Range)
    Dim ws As Worksheet
    Set ws = cell.Worksheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    cell.Select
    ActiveWindow.FreezePanes = True
End Sub

' 【FreezeTopRow】先頭行でウィンドウ枠を固定する。
' 使用例:
'   FreezeTopRow Worksheets("一覧")
' 解説: 「一覧」の 1 行目だけを固定する。下へスクロールしても見出しが残る。
Public Sub FreezeTopRow(ByVal ws As Worksheet)
    FreezePanesAt ws.Range("A2")
End Sub

' 【SetZoom】表示倍率を変える。10 から 400。
' 使用例:
'   SetZoom ActiveSheet, 90
' 解説: 今開いているシートの表示倍率を 90% にする。指定できるのは 10 から 400。
Public Sub SetZoom(ByVal ws As Worksheet, Optional ByVal percent As Long = 100)
    If percent < 10 Or percent > 400 Then Err.Raise 5, "SetZoom", "表示倍率は 10 から 400 までです。"
    ws.Activate
    ActiveWindow.Zoom = percent
End Sub

' 【ApplyNumberFormat】表示形式を設定する。書式文字列は VBA の NumberFormat。
' 使用例:
'   ApplyNumberFormat Range("C2:C100"), "0.0%"
' 解説: C2:C100 を、小数第 1 位までのパーセント表示にする。0.125 は 12.5% と見える。
Public Sub ApplyNumberFormat(ByVal target As Range, ByVal formatText As String)
    target.NumberFormat = formatText
End Sub

' 【FormatAsDate】yyyy-mm-dd の日付表示にする。
' 使用例:
'   FormatAsDate Worksheets("売上").Range("A2:A500")
' 解説: 「売上」の A2:A500 を、2026-09-22 のような日付表示にする。セルの値そのものは日付のまま。
Public Sub FormatAsDate(ByVal target As Range)
    target.NumberFormat = "yyyy-mm-dd"
End Sub

' 【FormatAsInteger】桁区切りの整数表示（#,##0）にする。
' 使用例:
'   FormatAsInteger Range("D2:D500")
' 解説: D2:D500 を、1234 が 1,234 と見える桁区切りの整数表示にする。
Public Sub FormatAsInteger(ByVal target As Range)
    target.NumberFormat = "#,##0"
End Sub

' 【FormatAsYen】円記号付きの桁区切り（¥#,##0）にする。
' 使用例:
'   FormatAsYen Range("E2:E500")
' 解説: E2:E500 を、1200 が ¥1,200 と見える円表示にする。
Public Sub FormatAsYen(ByVal target As Range)
    target.NumberFormat = """¥""#,##0"
End Sub

' 【SetupPrintPage】A4、余白 1.5cm 前後、見出し行を各ページに繰り返す印刷設定にする。
'   landscape が True なら横向き。fitWide と fitTall は横と縦のページ数。
'   プリンタと通信するため、初回は時間がかかることがある。
' 使用例:
'   SetupPrintPage Worksheets("一覧")                 ' 縦、1 ページ × 1 ページ
'   SetupPrintPage Worksheets("一覧"), True, 1, 1     ' 横向き
' 解説: 1 行目は「一覧」を A4 縦、横 1 ページ・縦 1 ページに収める。2 行目は同じ用紙で横向きにする。どちらも見出し行を各ページに繰り返す。
Public Sub SetupPrintPage(ByVal ws As Worksheet, _
    Optional ByVal landscape As Boolean = False, _
    Optional ByVal fitWide As Long = 1, _
    Optional ByVal fitTall As Long = 1)

    On Error GoTo EH
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = IIf(landscape, xlLandscape, xlPortrait)
        .Zoom = False
        .FitToPagesWide = fitWide
        .FitToPagesTall = fitTall
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(1.8)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
        .PrintTitleRows = "$1:$1"
        .CenterHeader = "&A"
        .RightHeader = "&D"
        .CenterFooter = "&P / &N"
    End With
    Exit Sub
EH:
    Err.Raise Err.Number, "SetupPrintPage", _
        "印刷設定を変更できませんでした。プリンタの状態を確認してください。" & vbCrLf & Err.Description
End Sub

' 【Run_AutoFitSelection】マクロ一覧用。選択範囲の列幅を内容に合わせ、幅 8〜40 に収める。
' 使用例:
'   表の範囲を選択してから実行する。
' 解説: 選んだ列の幅が内容に合い、狭すぎる列は幅 8、広すぎる列は幅 40 に収まる。
Public Sub Run_AutoFitSelection()
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    On Error GoTo EH
    AutoFitLimited Selection
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "AutoFitLimited"
End Sub

' 【Run_FreezeTopRow】マクロ一覧用。アクティブシートの先頭行でウィンドウ枠を固定する。
' 使用例:
'   見出しが 1 行目にあるシートを表示して実行する。
' 解説: 今開いているシートの 1 行目が固定され、下へスクロールしても見出しが残る。
Public Sub Run_FreezeTopRow()
    On Error GoTo EH
    FreezeTopRow ActiveSheet
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "FreezeTopRow"
End Sub

' 【Run_PrintSetupA4】マクロ一覧用。アクティブシートを A4・1 ページ幅の印刷設定にする。
' 使用例:
'   印刷したいシートを表示して実行する。終わるとメッセージが出る。
' 解説: 今のシートが A4・横 1 ページ・縦 1 ページの印刷設定になり、完了メッセージが出る。向きは縦。
Public Sub Run_PrintSetupA4()
    On Error GoTo EH
    SetupPrintPage ActiveSheet
    MsgBox "A4・1 ページ幅の印刷設定にしました。", vbInformation, "SetupPrintPage"
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "SetupPrintPage"
End Sub
