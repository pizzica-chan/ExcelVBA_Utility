Attribute VB_Name = "modArray"
Option Explicit

' セル範囲と二次元配列の読み書き。
' このモジュールだけでインポートして使えます。

' 【IsAllocated】配列に要素があるか。未初期化の動的配列は False。
' 使用例:
'   If IsAllocated(values) Then Debug.Print UBound(values)
' 解説: values に要素があれば、最後の添字をイミディエイトへ出す。まだ ReDim していない配列は要素が無いので、UBound でエラーにならない。
Public Function IsAllocated(ByVal values As Variant) As Boolean
    Dim lower As Long
    If Not IsArray(values) Then Exit Function
    On Error Resume Next
    lower = LBound(values)
    IsAllocated = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' 【ToMatrix】範囲を二次元配列で読む。1 セルでも 1 行 1 列の配列になる。
'   ループする前に一度読むと、セルを 1 つずつ触るより速い。
' 使用例:
'   data = ToMatrix(Worksheets("売上").Range("A2:C100"))
'   For r = 1 To UBound(data, 1)
'       Debug.Print data(r, 1), data(r, 2)
'   Next r
' 解説: 「売上」の A2:C100 を一度に配列へ読み、1 行ずつ 1 列目と 2 列目をイミディエイトへ出す。セルを 1 つずつ読むより速い。
Public Function ToMatrix(ByVal source As Range) As Variant
    Dim data() As Variant
    If source Is Nothing Then Err.Raise 5, "ToMatrix", "範囲が空です。"
    If source.Cells.CountLarge = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = source.Value
        ToMatrix = data
    Else
        ToMatrix = source.Value
    End If
End Function

' 【WriteMatrix】配列を左上セルからシートへ書く。1 次元配列は 1 行として書く。
' 使用例:
'   data = ToMatrix(Worksheets("元").Range("A1:C10"))
'   WriteMatrix data, Worksheets("先").Range("A1")
' 解説: 「元」の A1:C10 を配列で読み、「先」の A1 から同じ形で書き戻す。値だけが移る。
Public Sub WriteMatrix(ByVal values As Variant, ByVal topLeft As Range)
    Dim rows As Long
    Dim cols As Long
    If Not IsArray(values) Then
        topLeft.Cells(1, 1).Value = values
        Exit Sub
    End If
    If ArrayRank(values) = 1 Then
        rows = 1
        cols = UBound(values) - LBound(values) + 1
        topLeft.Cells(1, 1).Resize(rows, cols).Value = values
    Else
        rows = UBound(values, 1) - LBound(values, 1) + 1
        cols = UBound(values, 2) - LBound(values, 2) + 1
        topLeft.Cells(1, 1).Resize(rows, cols).Value = values
    End If
End Sub

' 【MatrixRows】配列の行数。配列でなければ 0。
' 使用例:
'   rowCount = MatrixRows(ToMatrix(Range("A1:C10")))   ' 10
' 解説: A1:C10 は 10 行なので、rowCount は 10。配列でない値を渡したときは 0。
Public Function MatrixRows(ByVal values As Variant) As Long
    If Not IsArray(values) Then Exit Function
    MatrixRows = UBound(values, 1) - LBound(values, 1) + 1
End Function

' 【MatrixCols】配列の列数。1 次元配列はその要素数。
' 使用例:
'   colCount = MatrixCols(data)
' 解説: data が二次元配列なら列数、1 次元配列なら要素数を colCount に入れる。配列でなければ 0。
Public Function MatrixCols(ByVal values As Variant) As Long
    If Not IsArray(values) Then Exit Function
    If ArrayRank(values) = 1 Then
        MatrixCols = UBound(values) - LBound(values) + 1
    Else
        MatrixCols = UBound(values, 2) - LBound(values, 2) + 1
    End If
End Function

' 【ColumnSlice】二次元配列から 1 列を切り出し、1 列の二次元配列で返す。
'   columnIndex は配列の添字そのもの。ToMatrix の結果なら 1 が左端。
' 使用例:
'   names = ColumnSlice(ToMatrix(Range("A1:C10")), 1)
'   WriteMatrix names, Range("E1")
' 解説: A1:C10 の左端の列だけを切り出し、E1 から下へ 1 列で書き出す。
Public Function ColumnSlice(ByVal values As Variant, ByVal columnIndex As Long) As Variant
    Dim rowIndex As Long
    Dim out() As Variant
    Dim lowerRow As Long
    Dim upperRow As Long
    If ArrayRank(values) <> 2 Then Err.Raise 5, "ColumnSlice", "二次元配列を指定してください。"
    lowerRow = LBound(values, 1)
    upperRow = UBound(values, 1)
    ReDim out(1 To upperRow - lowerRow + 1, 1 To 1)
    For rowIndex = lowerRow To upperRow
        out(rowIndex - lowerRow + 1, 1) = values(rowIndex, columnIndex)
    Next rowIndex
    ColumnSlice = out
End Function

' 【TransposeMatrix】行と列を入れ替える。
' 使用例:
'   turned = TransposeMatrix(ToMatrix(Range("A1:C2")))
'   WriteMatrix turned, Range("E1")
' 解説: A1:C2 は 2 行 3 列。行と列を入れ替えると 3 行 2 列になり、それを E1 から書く。
Public Function TransposeMatrix(ByVal values As Variant) As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim out() As Variant
    Dim lowerRow As Long
    Dim upperRow As Long
    Dim lowerCol As Long
    Dim upperCol As Long
    If ArrayRank(values) <> 2 Then Err.Raise 5, "TransposeMatrix", "二次元配列を指定してください。"
    lowerRow = LBound(values, 1)
    upperRow = UBound(values, 1)
    lowerCol = LBound(values, 2)
    upperCol = UBound(values, 2)
    ReDim out(1 To upperCol - lowerCol + 1, 1 To upperRow - lowerRow + 1)
    For rowIndex = lowerRow To upperRow
        For colIndex = lowerCol To upperCol
            out(colIndex - lowerCol + 1, rowIndex - lowerRow + 1) = values(rowIndex, colIndex)
        Next colIndex
    Next rowIndex
    TransposeMatrix = out
End Function

' 【WriteCollection】Collection の中身を、左上セルから下へ 1 列で書く。
' 使用例:
'   WriteCollection ListFiles(folder, "*.xlsx"), Range("A1"), "ファイル"
' 解説: folder 内の xlsx のパスを、A1 に見出し「ファイル」を置いて、A2 から下へ 1 列で書く。
Public Sub WriteCollection(ByVal items As Collection, ByVal topLeft As Range, Optional ByVal header As String = "")
    Dim rowIndex As Long
    Dim startRow As Long
    Dim item As Variant
    startRow = 1
    If Len(header) > 0 Then
        topLeft.Cells(1, 1).Value = header
        startRow = 2
    End If
    rowIndex = startRow
    For Each item In items
        topLeft.Cells(rowIndex, 1).Value = item
        rowIndex = rowIndex + 1
    Next item
End Sub

Private Function ArrayRank(ByVal values As Variant) As Long
    On Error Resume Next
    ArrayRank = 2
    Dim probe As Long
    probe = LBound(values, 2)
    If Err.Number <> 0 Then ArrayRank = 1
    Err.Clear
    On Error GoTo 0
End Function
