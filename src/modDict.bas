Attribute VB_Name = "modDict"
Option Explicit

' Scripting.Dictionary の作成と、セル範囲からの読み込み。
' 参照設定は不要です。このモジュールだけでインポートして使えます。
' compareMode は 0 が区別あり、1 が大文字小文字を区別しない比較です。
' 数値の 1 と文字列の "1" は別のキーです。

' 【CreateDictionary】参照設定なしで Dictionary を作る。
'   compareMode は 0 が大文字小文字を区別、1 が区別しない。既定は 1。
'   数値の 1 と文字列の "1" は、どちらでも別のキーになる。
' 使用例:
'   Set dict = CreateDictionary()
'   dict("東京") = "関東"
' 解説: 大文字小文字を区別しない辞書を作り、「東京」というキーに「関東」を入れる。あとから dict("東京") で「関東」を取り出せる。
Public Function CreateDictionary(Optional ByVal compareMode As Long = 1) As Object
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
    CreateDictionary.CompareMode = compareMode
End Function

' 【LoadDictionary】キー範囲と値範囲から辞書を作る。範囲の大きさは揃える。
'   同じキーが複数あるときは後のセルで上書きする。空のキーは入れない。
' 使用例:
'   ' A2:A100 がコード、B2:B100 が名称
'   Set master = LoadDictionary(Range("A2:A100"), Range("B2:B100"))
'   Debug.Print DictItem(master, "A001", "不明")
' 解説: A2:A100 のコードをキー、B2:B100 の名称を値にした辞書を作る。コード A001 の名称を出し、そのコードが無ければ「不明」と出す。
Public Function LoadDictionary(ByVal keyRange As Range, ByVal valueRange As Range, _
    Optional ByVal compareMode As Long = 1) As Object

    Dim keys As Variant
    Dim values As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim key As Variant
    Dim dict As Object
    If keyRange.Rows.Count <> valueRange.Rows.Count Or keyRange.Columns.Count <> valueRange.Columns.Count Then
        Err.Raise 5, "LoadDictionary", "キー範囲と値範囲の大きさが違います。"
    End If
    Set dict = CreateDictionary(compareMode)
    Set LoadDictionary = dict
    keys = ToMatrix(keyRange)
    values = ToMatrix(valueRange)
    For rowIndex = 1 To UBound(keys, 1)
        For colIndex = 1 To UBound(keys, 2)
            key = keys(rowIndex, colIndex)
            If Not IsBlankKey(key) Then dict(key) = values(rowIndex, colIndex)
        Next colIndex
    Next rowIndex
End Function

' 【DictExists】キーがあるか。
' 使用例:
'   If DictExists(master, Range("A2").Value) Then Range("B2").Value = master(Range("A2").Value)
' 解説: A2 のコードが辞書にあれば、対応する名称を B2 へ書く。無いときは B2 を変えない。
Public Function DictExists(ByVal dict As Object, ByVal key As Variant) As Boolean
    DictExists = dict.Exists(key)
End Function

' 【DictItem】キーの値。無いときは ifMissing。
' 使用例:
'   name = DictItem(master, "A001", "不明")
' 解説: 辞書からコード A001 の名称を取る。登録が無ければ name は「不明」。
Public Function DictItem(ByVal dict As Object, ByVal key As Variant, Optional ByVal ifMissing As Variant = Empty) As Variant
    If dict.Exists(key) Then
        DictItem = dict.Item(key)
    Else
        DictItem = ifMissing
    End If
End Function

' 【WriteLookup】キー範囲と同じ形で、辞書の値を destination へ書く。
'   見つからないセルには ifMissing を入れる。VLOOKUP のまとめて版。
' 使用例:
'   Set master = LoadDictionary(Sheets("マスタ").Range("A2:A100"), Sheets("マスタ").Range("B2:B100"))
'   WriteLookup Sheets("入力").Range("A2:A50"), master, Sheets("入力").Range("B2"), "該当なし"
' 解説: 「マスタ」のコードと名称で辞書を作り、「入力」の A2:A50 のコードに対応する名称を B2 から下へ書く。見つからない行は「該当なし」。
Public Sub WriteLookup(ByVal keyRange As Range, ByVal dict As Object, ByVal destination As Range, _
    Optional ByVal ifMissing As Variant = "")

    Dim keys As Variant
    Dim out() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim key As Variant
    keys = ToMatrix(keyRange)
    ReDim out(1 To UBound(keys, 1), 1 To UBound(keys, 2))
    For rowIndex = 1 To UBound(keys, 1)
        For colIndex = 1 To UBound(keys, 2)
            key = keys(rowIndex, colIndex)
            If dict.Exists(key) Then
                out(rowIndex, colIndex) = dict.Item(key)
            Else
                out(rowIndex, colIndex) = ifMissing
            End If
        Next colIndex
    Next rowIndex
    destination.Cells(1, 1).Resize(UBound(out, 1), UBound(out, 2)).Value = out
End Sub

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

Private Function IsBlankKey(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        IsBlankKey = True
        Exit Function
    End If
    If VarType(value) = vbString Then IsBlankKey = (Len(value) = 0)
End Function
