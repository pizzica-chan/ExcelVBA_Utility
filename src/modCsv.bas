Attribute VB_Name = "modCsv"
Option Explicit

' CSV の書き出しと読み込み。区切りはカンマ固定です。
' 文字コードは UTF-8（BOM 付き）または Shift_JIS を指定できます。
' このモジュールだけでインポートして使えます。
' 日付は yyyy-mm-dd、小数点はピリオドで書き出します。

' 【ExportRangeCsv】範囲を CSV にする。区切りはカンマ。引用符とセル内改行も扱う。
'   charset は "UTF-8"（BOM 付き）または "Shift_JIS"。
'   日付は yyyy-mm-dd、小数点はピリオド。
' 使用例:
'   ExportRangeCsv Worksheets("売上").Range("A1:F200"), "C:\temp\売上.csv"
'   ExportRangeCsv Selection, CombinePath(DesktopFolder(), "out.csv"), "Shift_JIS"
' 解説: 1 行目は「売上」の A1:F200 を、C:\temp\売上.csv へ UTF-8 で書き出す。2 行目は選択範囲を、デスクトップの out.csv へ Shift_JIS で書き出す。
Public Sub ExportRangeCsv(ByVal source As Range, ByVal filePath As String, Optional ByVal charset As String = "UTF-8")
    Dim data As Variant
    Dim lines() As String
    Dim fields() As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowCount As Long
    Dim colCount As Long
    If source Is Nothing Then Err.Raise 5, "ExportRangeCsv", "範囲が空です。"
    If source.Areas.Count <> 1 Then Err.Raise 5, "ExportRangeCsv", "連続した 1 つの範囲を指定してください。"
    data = ToMatrix(source)
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    ReDim lines(1 To rowCount)
    For rowIndex = 1 To rowCount
        ReDim fields(1 To colCount)
        For colIndex = 1 To colCount
            fields(colIndex) = VariantToCsvField(data(rowIndex, colIndex))
        Next colIndex
        lines(rowIndex) = Join(fields, ",")
    Next rowIndex
    WriteTextFile filePath, Join(lines, vbCrLf) & vbCrLf, charset
End Sub

' 【ExportSheetCsv】シートの使用範囲を CSV にする。データが無ければ空ファイル。
' 使用例:
'   ExportSheetCsv Worksheets("売上"), CombinePath(ThisFolder(), "売上.csv")
' 解説: 「売上」シートで使っている範囲を、マクロのブックと同じフォルダの売上.csv にする。データが無ければ空のファイルになる。
Public Sub ExportSheetCsv(ByVal ws As Worksheet, ByVal filePath As String, Optional ByVal charset As String = "UTF-8")
    Dim source As Range
    Set source = UsedDataRange(ws)
    If source Is Nothing Then
        WriteTextFile filePath, "", charset
    Else
        ExportRangeCsv source, filePath, charset
    End If
End Sub

' 【ExportAllSheets】表示中の各シートを、シート名.csv としてフォルダへ書き出す。
'   フォルダが無ければ作る。非表示シートは対象外。
' 使用例:
'   ExportAllSheets ThisWorkbook, CombinePath(ThisFolder(), "csv")
' 解説: このブックの表示中シートを、マクロと同じ場所の csv フォルダへ、シート名.csv として書き出す。フォルダが無ければ作る。
Public Sub ExportAllSheets(ByVal wb As Workbook, ByVal folderPath As String, Optional ByVal charset As String = "UTF-8")
    Dim ws As Worksheet
    EnsureFolderPath folderPath
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ExportSheetCsv ws, CombinePath(folderPath, SafeFileStem(ws.Name) & ".csv"), charset
        End If
    Next ws
End Sub

' 【ImportCsv】CSV を destination の左上から書き込む。
'   書き込んだ範囲より下に残っている古い値は消さない。先に範囲を消すこと。
'   文字コードは書き出したときと同じものを指定する。
' 使用例:
'   Worksheets("取込").Cells.ClearContents
'   ImportCsv "C:\temp\売上.csv", Worksheets("取込").Range("A1")
' 解説: 先に「取込」シートの中身を消し、その A1 から C:\temp\売上.csv を UTF-8 として読み込む。消さずに読むと、新しい表より下の古い行が残る。
Public Sub ImportCsv(ByVal filePath As String, ByVal destination As Range, Optional ByVal charset As String = "UTF-8")
    Dim content As String
    Dim rows As Collection
    Dim rowValues As Variant
    Dim data() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowCount As Long
    Dim colCount As Long
    If Not FileExists(filePath) Then Err.Raise 53, "ImportCsv", "ファイルが見つかりません: " & filePath
    content = ReadTextFile(filePath, charset)
    Set rows = ParseCsv(content)
    rowCount = rows.Count
    If rowCount = 0 Then Exit Sub
    For rowIndex = 1 To rowCount
        rowValues = rows(rowIndex)
        If UBound(rowValues) > colCount Then colCount = UBound(rowValues)
    Next rowIndex
    ReDim data(1 To rowCount, 1 To colCount)
    For rowIndex = 1 To rowCount
        rowValues = rows(rowIndex)
        For colIndex = 1 To UBound(rowValues)
            data(rowIndex, colIndex) = CsvFieldToValue(CStr(rowValues(colIndex)))
        Next colIndex
    Next rowIndex
    destination.Cells(1, 1).Resize(rowCount, colCount).Value = data
End Sub

' 【Run_ExportSelectionToCsv】マクロ一覧用。選択範囲を UTF-8 の CSV として保存する。
'   保存ダイアログで場所を選ぶ。キャンセルすると何もしない。
' 使用例:
'   書き出したい表を選択してから実行する。
' 解説: 保存ダイアログで場所とファイル名を選ぶと、選択範囲が UTF-8 の CSV になる。キャンセルすると何もしない。
Public Sub Run_ExportSelectionToCsv()
    Dim filePath As String
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    filePath = PickSaveCsv()
    If Len(filePath) = 0 Then Exit Sub
    On Error GoTo EH
    ExportRangeCsv Selection, filePath, "UTF-8"
    MsgBox "書き出しました。" & vbCrLf & filePath, vbInformation, "ExportRangeCsv"
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "ExportRangeCsv"
End Sub

' 【Run_ImportCsvToSelection】マクロ一覧用。CSV を選択セルの位置から読み込む。
'   読み込み先より下の古いデータは残るので、空のセルを選ぶ。
' 使用例:
'   貼り付け先の左上セルを選択してから実行し、UTF-8 の CSV を選ぶ。
' 解説: 選んだセルを左上にして CSV を貼り付ける。その下に古いデータが残るので、空のセルか、先に消した範囲を選ぶ。
Public Sub Run_ImportCsvToSelection()
    Dim filePath As String
    If TypeName(Selection) <> "Range" Then
        MsgBox "貼り付け先のセルを選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    filePath = PickOpenCsv()
    If Len(filePath) = 0 Then Exit Sub
    On Error GoTo EH
    ImportCsv filePath, Selection.Cells(1, 1), "UTF-8"
    MsgBox "読み込みました。", vbInformation, "ImportCsv"
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "ImportCsv"
End Sub

Private Function PickSaveCsv() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    With dlg
        .Title = "CSV の保存先"
        .InitialFileName = "export.csv"
        .Filters.Clear
        .Filters.Add "CSV ファイル", "*.csv"
        If .Show <> -1 Then Exit Function
        PickSaveCsv = .SelectedItems(1)
    End With
End Function

Private Function PickOpenCsv() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "CSV を選択"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV ファイル", "*.csv"
        .Filters.Add "すべてのファイル", "*.*"
        If .Show <> -1 Then Exit Function
        PickOpenCsv = .SelectedItems(1)
    End With
End Function

Private Function VariantToCsvField(ByVal value As Variant) As String
    Dim text As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarType(value) = vbDate Then
        text = Format$(value, "yyyy-mm-dd")
        If TimeValue(value) <> 0 Then text = text & " " & Format$(value, "hh:nn:ss")
    ElseIf VarType(value) = vbDouble Or VarType(value) = vbSingle Or VarType(value) = vbCurrency Then
        text = Trim$(Str$(value))
    ElseIf VarType(value) = vbBoolean Then
        text = IIf(value, "TRUE", "FALSE")
    Else
        text = CStr(value)
    End If
    VariantToCsvField = QuoteCsv(text)
End Function

Private Function QuoteCsv(ByVal text As String) As String
    Dim needsQuote As Boolean
    needsQuote = (InStr(text, ",") > 0) Or (InStr(text, """") > 0) _
        Or (InStr(text, vbLf) > 0) Or (InStr(text, vbCr) > 0)
    text = Replace(text, """", """""")
    If needsQuote Then
        QuoteCsv = """" & text & """"
    Else
        QuoteCsv = text
    End If
End Function

Private Function CsvFieldToValue(ByVal text As String) As Variant
    If Len(text) = 0 Then
        CsvFieldToValue = Empty
        Exit Function
    End If
    If text Like "####-##-##" Then
        CsvFieldToValue = DateSerial(CLng(Left$(text, 4)), CLng(Mid$(text, 6, 2)), CLng(Right$(text, 2)))
        Exit Function
    End If
    If text Like "####-##-## ##:##:##" Then
        CsvFieldToValue = DateSerial(CLng(Left$(text, 4)), CLng(Mid$(text, 6, 2)), CLng(Mid$(text, 9, 2))) _
            + TimeSerial(CLng(Mid$(text, 12, 2)), CLng(Mid$(text, 15, 2)), CLng(Right$(text, 2)))
        Exit Function
    End If
    If IsPlainNumber(text) Then
        CsvFieldToValue = Val(text)
        Exit Function
    End If
    CsvFieldToValue = text
End Function

Private Function IsPlainNumber(ByVal text As String) As Boolean
    Dim body As String
    Dim i As Long
    Dim ch As String
    Dim dots As Long
    body = text
    If Left$(body, 1) = "+" Or Left$(body, 1) = "-" Then body = Mid$(body, 2)
    If Len(body) = 0 Then Exit Function
    If Left$(body, 1) = "0" And body <> "0" And Left$(body, 2) <> "0." Then Exit Function
    For i = 1 To Len(body)
        ch = Mid$(body, i, 1)
        If ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function
        End If
    Next i
    IsPlainNumber = True
End Function

Private Function ParseCsv(ByVal content As String) As Collection
    Dim rows As New Collection
    Dim fields As New Collection
    Dim i As Long
    Dim charCount As Long
    Dim ch As String
    Dim field As String
    Dim inQuotes As Boolean
    Dim pendingRow As Boolean
    If Len(content) > 0 Then
        If AscW(Left$(content, 1)) = &HFEFF Then content = Mid$(content, 2)
    End If
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    charCount = Len(content)
    For i = 1 To charCount
        ch = Mid$(content, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < charCount And Mid$(content, i + 1, 1) = """" Then
                    field = field & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                field = field & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQuotes = True
                    pendingRow = True
                Case ","
                    fields.Add field
                    field = ""
                    pendingRow = True
                Case vbLf
                    fields.Add field
                    field = ""
                    rows.Add CollectionToArray(fields)
                    Set fields = New Collection
                    pendingRow = False
                Case Else
                    field = field & ch
                    pendingRow = True
            End Select
        End If
    Next i
    If pendingRow Or fields.Count > 0 Then
        fields.Add field
        rows.Add CollectionToArray(fields)
    End If
    Set ParseCsv = rows
End Function

Private Function CollectionToArray(ByVal fields As Collection) As Variant
    Dim values() As String
    Dim i As Long
    If fields.Count = 0 Then
        ReDim values(1 To 1)
        CollectionToArray = values
        Exit Function
    End If
    ReDim values(1 To fields.Count)
    For i = 1 To fields.Count
        values(i) = CStr(fields(i))
    Next i
    CollectionToArray = values
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String, ByVal charset As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = CanonicalCharset(charset)
    stream.Open
    If Len(content) > 0 Then stream.WriteText content
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Function ReadTextFile(ByVal filePath As String, ByVal charset As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = CanonicalCharset(charset)
    stream.Open
    stream.LoadFromFile filePath
    ReadTextFile = stream.ReadText
    stream.Close
End Function

Private Function CanonicalCharset(ByVal charset As String) As String
    If StrComp(charset, "utf-8", vbTextCompare) = 0 Or StrComp(charset, "utf8", vbTextCompare) = 0 Then
        CanonicalCharset = "UTF-8"
    ElseIf StrComp(charset, "shift_jis", vbTextCompare) = 0 Or StrComp(charset, "sjis", vbTextCompare) = 0 Then
        CanonicalCharset = "Shift_JIS"
    Else
        CanonicalCharset = charset
    End If
End Function

Private Function ToMatrix(ByVal source As Range) As Variant
    Dim data() As Variant
    If source.Cells.CountLarge = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = source.Value
        ToMatrix = data
    Else
        ToMatrix = source.Value
    End If
End Function

Private Function UsedDataRange(ByVal ws As Worksheet) As Range
    Dim lastCell As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Exit Function
    rowIndex = lastCell.Row
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    colIndex = lastCell.Column
    Set UsedDataRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowIndex, colIndex))
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(filePath) > 0 And Len(Dir$(filePath, vbNormal)) > 0)
    On Error GoTo 0
End Function

Private Sub EnsureFolderPath(ByVal folderPath As String)
    Dim parentPath As String
    Dim slash As Long
    folderPath = Trim$(folderPath)
    Do While Len(folderPath) > 3 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop
    If Len(folderPath) = 0 Then Err.Raise 5, "ExportAllSheets", "フォルダが指定されていません。"
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub
    slash = InStrRev(folderPath, "\")
    If slash > 0 Then
        parentPath = Left$(folderPath, slash - 1)
        If Len(parentPath) = 2 And Mid$(parentPath, 2, 1) = ":" Then parentPath = parentPath & "\"
        If Len(parentPath) > 3 Then EnsureFolderPath parentPath
    End If
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function CombinePath(ByVal folderPath As String, ByVal name As String) As String
    If Right$(folderPath, 1) = "\" Then
        CombinePath = folderPath & name
    Else
        CombinePath = folderPath & "\" & name
    End If
End Function

Private Function SafeFileStem(ByVal name As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        name = Replace(name, CStr(ch), "_")
    Next ch
    name = Trim$(name)
    If Len(name) = 0 Then name = "sheet"
    SafeFileStem = name
End Function
