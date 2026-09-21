Attribute VB_Name = "modAudit"
Option Explicit

' シート、名前定義、外部リンク、ハイパーリンク、数式の一覧。
' このモジュールだけでインポートして使えます。
' 各 Write 関数は、書き込んだ行数（見出しを含む）を返します。

' 【WriteSheetIndex】シート名、表示状態、使用範囲を destination へ書く。
'   戻り値は書き込んだ行数（見出しを含む）。次の一覧を続ける位置の計算に使う。
' 使用例:
'   nextRow = 1 + WriteSheetIndex(ThisWorkbook, Worksheets("点検").Range("A1"))
' 解説: このブックのシート名、表示状態、使用範囲を「点検」の A1 から書く。戻り値は見出しを含む行数なので、1 を足すと次の一覧を書き始める行になる。
Public Function WriteSheetIndex(ByVal wb As Workbook, ByVal destination As Range) As Long
    Dim ws As Worksheet
    Dim rowIndex As Long
    WriteHeader destination, Array("シート名", "表示", "使用範囲")
    rowIndex = 2
    For Each ws In wb.Worksheets
        destination.Cells(rowIndex, 1).Value = ws.Name
        destination.Cells(rowIndex, 2).Value = VisibilityText(ws.Visible)
        destination.Cells(rowIndex, 3).Value = ws.UsedRange.Address(False, False)
        rowIndex = rowIndex + 1
    Next ws
    WriteSheetIndex = rowIndex - 1
End Function

' 【WriteDefinedNames】名前定義の名前、参照先、表示状態を書く。
' 使用例:
'   WriteDefinedNames ThisWorkbook, Worksheets("点検").Range("A20")
' 解説: このブックの名前定義を、「点検」の A20 から名前・参照先・表示状態の 3 列で書く。
Public Function WriteDefinedNames(ByVal wb As Workbook, ByVal destination As Range) As Long
    Dim definedName As Name
    Dim refersTo As String
    Dim rowIndex As Long
    WriteHeader destination, Array("名前", "参照先", "表示")
    rowIndex = 2
    For Each definedName In wb.Names
        refersTo = ""
        On Error Resume Next
        refersTo = definedName.RefersTo
        If Err.Number <> 0 Then refersTo = "(参照を取得できません)"
        Err.Clear
        On Error GoTo 0
        destination.Cells(rowIndex, 1).Value = definedName.Name
        destination.Cells(rowIndex, 2).NumberFormat = "@"
        destination.Cells(rowIndex, 2).Value = refersTo
        destination.Cells(rowIndex, 3).Value = IIf(definedName.Visible, "表示", "非表示")
        rowIndex = rowIndex + 1
    Next definedName
    If rowIndex = 2 Then
        destination.Cells(2, 1).Value = "（なし）"
        rowIndex = 3
    End If
    WriteDefinedNames = rowIndex - 1
End Function

' 【WriteExternalLinks】他ブックへのリンク一覧を書く。無ければ「（なし）」。
' 使用例:
'   WriteExternalLinks ThisWorkbook, Worksheets("点検").Range("A40")
' 解説: このブックが参照している他ブックを、「点検」の A40 から一覧にする。リンクが無ければ「（なし）」と書く。
Public Function WriteExternalLinks(ByVal wb As Workbook, ByVal destination As Range) As Long
    Dim links As Variant
    Dim i As Long
    Dim rowIndex As Long
    WriteHeader destination, Array("外部リンク")
    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0
    rowIndex = 2
    If IsEmpty(links) Then
        destination.Cells(2, 1).Value = "（なし）"
        rowIndex = 3
    Else
        For i = LBound(links) To UBound(links)
            destination.Cells(rowIndex, 1).Value = links(i)
            rowIndex = rowIndex + 1
        Next i
    End If
    WriteExternalLinks = rowIndex - 1
End Function

' 【WriteHyperlinks】全シートのハイパーリンクを、シート・セル・表示文字・リンク先で書く。
' 使用例:
'   WriteHyperlinks ThisWorkbook, Worksheets("点検").Range("A50")
' 解説: 全シートのハイパーリンクを、「点検」の A50 からシート・セル・表示文字・リンク先の 4 列で書く。
Public Function WriteHyperlinks(ByVal wb As Workbook, ByVal destination As Range) As Long
    Dim ws As Worksheet
    Dim link As Hyperlink
    Dim rowIndex As Long
    WriteHeader destination, Array("シート", "セル", "表示文字", "リンク先")
    rowIndex = 2
    For Each ws In wb.Worksheets
        For Each link In ws.Hyperlinks
            destination.Cells(rowIndex, 1).Value = ws.Name
            destination.Cells(rowIndex, 2).Value = link.Range.Address(False, False)
            destination.Cells(rowIndex, 3).Value = link.TextToDisplay
            destination.Cells(rowIndex, 4).Value = link.Address
            rowIndex = rowIndex + 1
        Next link
    Next ws
    If rowIndex = 2 Then
        destination.Cells(2, 1).Value = "（なし）"
        rowIndex = 3
    End If
    WriteHyperlinks = rowIndex - 1
End Function

' 【WriteFormulaList】範囲内の数式を、セル・数式・値で書く。先頭 5000 件まで。
'   数式はテキストとして書くので、点検シート側で再計算されない。
' 使用例:
'   WriteFormulaList Worksheets("集計").UsedRange, Worksheets("点検").Range("A70")
' 解説: 「集計」で使っている範囲の数式を、「点検」の A70 からセル・数式・値で書く。数式は文字として書くので、点検シート側では再計算されない。先頭 5000 件まで。
Public Function WriteFormulaList(ByVal source As Range, ByVal destination As Range) As Long
    Dim formulas As Range
    Dim area As Range
    Dim cell As Range
    Dim rowIndex As Long
    Dim written As Long
    Const MaxItems As Long = 5000
    WriteHeader destination, Array("セル", "数式", "値")
    rowIndex = 2
    If Not source Is Nothing Then
        On Error Resume Next
        Set formulas = source.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
    End If
    If Not formulas Is Nothing Then
        For Each area In formulas.Areas
            For Each cell In area.Cells
                destination.Cells(rowIndex, 1).Value = cell.Address(False, False)
                destination.Cells(rowIndex, 2).Value = "'" & cell.Formula
                If IsError(cell.Value) Then
                    destination.Cells(rowIndex, 3).Value = cell.Text
                Else
                    destination.Cells(rowIndex, 3).Value = cell.Value
                End If
                rowIndex = rowIndex + 1
                written = written + 1
                If written >= MaxItems Then Exit For
            Next cell
            If written >= MaxItems Then Exit For
        Next area
        If written >= MaxItems Then destination.Cells(1, 4).Value = "先頭 " & CStr(MaxItems) & " 件まで"
    End If
    If rowIndex = 2 Then
        destination.Cells(2, 1).Value = "（なし）"
        rowIndex = 3
    End If
    WriteFormulaList = rowIndex - 1
End Function

' 【Run_WriteAuditSheets】マクロ一覧用。開いているブックの点検シートを作る。
'   シート一覧、名前定義、外部リンク、ハイパーリンクに加え、
'   実行時に開いていたシートの数式を「点検」シートへまとめる。
'   「点検」シートを表示したまま実行すると止まる。
' 使用例:
'   数式を確認したいシートを表示してから実行する。
' 解説: 「点検」シートが作られ、シート一覧、名前定義、外部リンク、ハイパーリンクと、実行前に開いていたシートの数式がまとまる。「点検」を表示したまま実行すると止まる。
Public Sub Run_WriteAuditSheets()
    Dim wb As Workbook
    Dim source As Worksheet
    Dim audit As Worksheet
    Dim rowIndex As Long
    Dim sourceRange As Range
    On Error GoTo EH
    Set wb = ActiveWorkbook
    Set source = ActiveSheet
    Set audit = EnsureAuditSheet(wb)
    If source Is audit Then
        MsgBox "点検シート以外を表示してから実行してください。数式一覧は、実行時に開いていたシートを対象にします。", vbExclamation
        Exit Sub
    End If
    Set sourceRange = UsedDataRange(source)
    audit.Cells.Clear
    audit.Range("A1").Value = "ブック点検"
    audit.Range("A1").Font.Bold = True
    audit.Range("A2").Value = "数式一覧の対象シート: " & source.Name
    rowIndex = 4
    rowIndex = rowIndex + WriteSheetIndex(wb, audit.Cells(rowIndex, 1)) + 1
    rowIndex = rowIndex + WriteDefinedNames(wb, audit.Cells(rowIndex, 1)) + 1
    rowIndex = rowIndex + WriteExternalLinks(wb, audit.Cells(rowIndex, 1)) + 1
    rowIndex = rowIndex + WriteHyperlinks(wb, audit.Cells(rowIndex, 1)) + 1
    WriteFormulaList sourceRange, audit.Cells(rowIndex, 1)
    audit.Columns("A:D").AutoFit
    audit.Activate
    audit.Range("A1").Select
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "ブック点検"
End Sub

Private Function EnsureAuditSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set EnsureAuditSheet = wb.Worksheets("点検")
    On Error GoTo 0
    If EnsureAuditSheet Is Nothing Then
        Set EnsureAuditSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureAuditSheet.Name = "点検"
    End If
End Function

Private Sub WriteHeader(ByVal destination As Range, ByVal labels As Variant)
    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        destination.Cells(1, i - LBound(labels) + 1).Value = labels(i)
    Next i
    destination.Resize(1, UBound(labels) - LBound(labels) + 1).Font.Bold = True
End Sub

Private Function VisibilityText(ByVal visibleState As XlSheetVisibility) As String
    Select Case visibleState
        Case xlSheetVisible: VisibilityText = "表示"
        Case xlSheetHidden: VisibilityText = "非表示"
        Case xlSheetVeryHidden: VisibilityText = "非常に非表示"
        Case Else: VisibilityText = CStr(visibleState)
    End Select
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
