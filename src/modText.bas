Attribute VB_Name = "modText"
Option Explicit

' 文字列の整形、全角半角、正規表現。
' このモジュールだけでインポートして使えます。
' NormalizeText、DigitsOnly、ToHalfWidth、ToHiragana、ToKatakana、PadLeft はワークシート関数としても使えます。

' 【Nz】空欄、Null、空文字のとき ifBlank を返す。0 は空欄にしない。
' 使用例:
'   name = Nz(Range("A2").Value, "（無名）")
' 解説: A2 が空なら name は「（無名）」。A2 に名前が入っていれば、その名前をそのまま使う。0 は空ではないので、0 のまま残る。
' ワークシート関数としても使える。=Nz(A2,"（無名）")
Public Function Nz(ByVal value As Variant, Optional ByVal ifBlank As Variant = "") As Variant
    If IsBlank(value) Then
        Nz = ifBlank
    Else
        Nz = value
    End If
End Function

' 【IsBlank】空欄、Null、空文字なら True。0 やエラー値は False。
' 使用例:
'   If IsBlank(Range("A2").Value) Then Range("A2").Value = "未入力"
' 解説: A2 が空欄なら「未入力」と書き込む。0 やエラー値が入っているときは空欄とみなさず、そのまま残す。
Public Function IsBlank(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsEmpty(value) Or IsNull(value) Then
        IsBlank = True
        Exit Function
    End If
    If VarType(value) = vbString Then IsBlank = (Len(value) = 0)
End Function

' 【Coalesce】左から見て、最初の非空白を返す。全部空なら空文字。
' 使用例:
'   label = Coalesce(Range("B2").Value, Range("A2").Value, "不明")
' 解説: B2 に値があればそれを使い、B2 が空なら A2 を使う。両方空なら「不明」になる。
Public Function Coalesce(ParamArray values() As Variant) As Variant
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Not IsBlank(values(i)) Then
            Coalesce = values(i)
            Exit Function
        End If
    Next i
    Coalesce = ""
End Function

' 【NormalizeText】前後の空白を除き、NBSP と全角スペースを半角スペースにする。
'   collapseSpaces が True のときは、連続する半角スペースも 1 つにする。
' 使用例:
'   Range("A2").Value = NormalizeText(Range("A2").Value, True)
' 解説: A2 の前後の空白を除き、全角スペースや連続スペースを半角 1 つに直して、同じセルへ書き戻す。
' ワークシート関数としても使える。=NormalizeText(A2,TRUE)
Public Function NormalizeText(ByVal value As Variant, Optional ByVal collapseSpaces As Boolean = False) As String
    Dim text As String
    If IsBlank(value) Or IsError(value) Then Exit Function
    text = CStr(value)
    text = Replace(text, ChrW(&HA0), " ")
    text = Replace(text, ChrW(&H3000), " ")
    text = Replace(text, vbTab, " ")
    text = TrimEdges(text)
    If collapseSpaces Then text = CollapseSpaceRuns(text)
    NormalizeText = text
End Function

' 【CollapseSpaces】前後を整え、連続スペースを 1 つにする。NormalizeText の省略形。
' 使用例:
'   Debug.Print CollapseSpaces("  東京　　支店  ")   ' 東京 支店
' 解説: 前後の空白を除き、全角スペースも含めて連続した空白を半角 1 つにする。結果は「東京 支店」。
Public Function CollapseSpaces(ByVal text As String) As String
    CollapseSpaces = NormalizeText(text, True)
End Function

' 【DigitsOnly】全角数字を半角にしてから、数字以外を除く。
' 使用例:
'   tel = DigitsOnly("TEL０３-1234-5678")   ' 0312345678
' 解説: 全角の０３を半角にし、TEL とハイフンを除く。残るのは数字の 0312345678。
' ワークシート関数としても使える。=DigitsOnly(A2)
Public Function DigitsOnly(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    text = StrConv(text, vbNarrow)
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

' 【ToHalfWidth】全角英数カナを半角にする。日本語環境向け。
' 使用例:
'   code = ToHalfWidth(Range("A2").Value)   ' Ａ１ → A1
' 解説: A2 の全角英数カナを半角にする。Ａ１ は A1 になる。漢字やひらがなはそのまま。
' ワークシート関数としても使える。=ToHalfWidth(A2)
Public Function ToHalfWidth(ByVal text As String) As String
    ToHalfWidth = StrConv(text, vbNarrow)
End Function

' 【ToFullWidth】半角英数カナを全角にする。
' 使用例:
'   Debug.Print ToFullWidth("ABC")   ' ＡＢＣ
' 解説: 半角の ABC を全角のＡＢＣにしてイミディエイトへ出す。
Public Function ToFullWidth(ByVal text As String) As String
    ToFullWidth = StrConv(text, vbWide)
End Function

' 【ToHiragana】カタカナをひらがなにする。
' 使用例:
'   Debug.Print ToHiragana("トウキョウ")   ' とうきょう
' 解説: カタカナのトウキョウを、ひらがなのとうきょうにして出す。
' ワークシート関数としても使える。=ToHiragana(A2)
Public Function ToHiragana(ByVal text As String) As String
    ToHiragana = StrConv(text, vbHiragana)
End Function

' 【ToKatakana】ひらがなをカタカナにする。
' 使用例:
'   Debug.Print ToKatakana("とうきょう")   ' トウキョウ
' 解説: ひらがなのとうきょうを、カタカナのトウキョウにして出す。
' ワークシート関数としても使える。=ToKatakana(A2)
Public Function ToKatakana(ByVal text As String) As String
    ToKatakana = StrConv(text, vbKatakana)
End Function

' 【PadLeft】文字数が width になるまで、左側を padChar で埋める。
' 使用例:
'   code = PadLeft(Range("A2").Value, 5, "0")   ' 7 → 00007
' 解説: A2 が 7 なら、5 桁になるまで左を 0 で埋め、00007 にする。すでに 5 文字以上ならそのまま。
' ワークシート関数としても使える。=PadLeft(A2,5,"0")
Public Function PadLeft(ByVal text As String, ByVal width As Long, Optional ByVal padChar As String = " ") As String
    Dim pad As String
    pad = PadCharacter(padChar)
    text = CStr(text)
    Do While Len(text) < width
        text = pad & text
    Loop
    PadLeft = text
End Function

' 【PadRight】文字数が width になるまで、右側を padChar で埋める。
' 使用例:
'   label = PadRight("氏名", 8, "　")
' 解説: 「氏名」の右側に全角スペースを足し、8 文字にする。帳票で見出しの幅を揃えたいときに使う。
Public Function PadRight(ByVal text As String, ByVal width As Long, Optional ByVal padChar As String = " ") As String
    Dim pad As String
    pad = PadCharacter(padChar)
    text = CStr(text)
    Do While Len(text) < width
        text = text & pad
    Loop
    PadRight = text
End Function

' 【PadLeftB】Shift_JIS 換算のバイト幅まで、半角 1 文字で左を埋める。固定長出力向け。
'   日本語 Windows では「あ」は 2 バイト。詰める文字は半角 1 文字にする。
'   幅は LenB ではなく、システムの ANSI（日本語 Windows では Shift_JIS）で数える。
' 使用例:
'   line = PadRightB(name, 20) & PadLeftB(amount, 10)
' 解説: 名前を右に空白で埋めて 20 バイト、金額を左に空白で埋めて 10 バイトにし、1 行の固定長データとしてつなぐ。
Public Function PadLeftB(ByVal text As String, ByVal byteWidth As Long, Optional ByVal padChar As String = " ") As String
    Dim pad As String
    pad = SingleBytePad(padChar)
    Do While ByteLen(text) < byteWidth
        text = pad & text
    Loop
    PadLeftB = text
End Function

' 【PadRightB】Shift_JIS 換算のバイト幅まで、半角 1 文字で右を埋める。
' 使用例:
'   fixedName = PadRightB(Range("A2").Value, 20, " ")
' 解説: A2 の文字を、Shift_JIS で 20 バイトになるまで右側に半角スペースを足す。「あ」は 2 バイトとして数える。
Public Function PadRightB(ByVal text As String, ByVal byteWidth As Long, Optional ByVal padChar As String = " ") As String
    Dim pad As String
    pad = SingleBytePad(padChar)
    Do While ByteLen(text) < byteWidth
        text = text & pad
    Loop
    PadRightB = text
End Function

' 【CountText】part が何回含まれるか。重なりは数えない（"aaa" の中の "aa" は 1 回）。
'   既定では大文字小文字を区別しない。
' 使用例:
'   n = CountText(Range("A2").Value, "株式会社")
' 解説: A2 の文章に「株式会社」が何回入っているかを n に入れる。大文字小文字は区別しない。
Public Function CountText(ByVal text As String, ByVal part As String, _
    Optional ByVal compareMethod As VbCompareMethod = vbTextCompare) As Long

    Dim startAt As Long
    Dim foundAt As Long
    If Len(part) = 0 Then Exit Function
    startAt = 1
    Do
        foundAt = InStr(startAt, text, part, compareMethod)
        If foundAt = 0 Then Exit Do
        CountText = CountText + 1
        startAt = foundAt + Len(part)
    Loop
End Function

' 【RegexFirst】正規表現に最初に一致した文字列。無ければ空文字。
' 使用例:
'   postal = RegexFirst(Range("A2").Value, "\d{3}-?\d{4}")
' 解説: A2 から最初の郵便番号を取り出す。123-4567 でも 1234567 でも一致する。無ければ postal は空文字。
Public Function RegexFirst(ByVal text As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = True) As String
    Dim matches As Object
    Set matches = RegexEngine(pattern, ignoreCase, False).Execute(text)
    If matches.Count > 0 Then RegexFirst = CStr(matches(0).Value)
End Function

' 【RegexMatch】正規表現に 1 つでも一致すれば True。
' 使用例:
'   If RegexMatch(mail, "^[^@]+@[^@]+$") Then Range("B2").Value = "OK"
' 解説: mail が「@ が 1 つある文字列」なら、B2 に OK と書く。@ が無い、または 2 つ以上あるときは何も書かない。
Public Function RegexMatch(ByVal text As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = True) As Boolean
    RegexMatch = RegexEngine(pattern, ignoreCase, False).Test(text)
End Function

' 【RegexReplace】正規表現に一致した箇所をすべて replacement に置き換える。
' 使用例:
'   cleaned = RegexReplace(Range("A2").Value, "\s+", " ")
' 解説: A2 の連続する空白や改行を、半角スペース 1 つに置き換えた文字列を cleaned に入れる。元のセルは変えない。
Public Function RegexReplace(ByVal text As String, ByVal pattern As String, ByVal replacement As String, _
    Optional ByVal ignoreCase As Boolean = True) As String

    RegexReplace = RegexEngine(pattern, ignoreCase, True).Replace(text, replacement)
End Function

Private Function RegexEngine(ByVal pattern As String, ByVal ignoreCase As Boolean, ByVal isGlobal As Boolean) As Object
    Set RegexEngine = CreateObject("VBScript.RegExp")
    RegexEngine.Pattern = pattern
    RegexEngine.IgnoreCase = ignoreCase
    RegexEngine.Global = isGlobal
End Function

Private Function TrimEdges(ByVal text As String) As String
    Dim ch As String
    text = Trim$(text)
    Do While Len(text) > 0
        ch = Left$(text, 1)
        If ch <> vbCr And ch <> vbLf And ch <> " " Then Exit Do
        text = Mid$(text, 2)
    Loop
    Do While Len(text) > 0
        ch = Right$(text, 1)
        If ch <> vbCr And ch <> vbLf And ch <> " " Then Exit Do
        text = Left$(text, Len(text) - 1)
    Loop
    TrimEdges = text
End Function

Private Function CollapseSpaceRuns(ByVal text As String) As String
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    CollapseSpaceRuns = text
End Function

Private Function PadCharacter(ByVal padChar As String) As String
    If Len(padChar) = 0 Then
        PadCharacter = " "
    Else
        PadCharacter = Left$(padChar, 1)
    End If
End Function

Private Function SingleBytePad(ByVal padChar As String) As String
    Dim pad As String
    pad = PadCharacter(padChar)
    If ByteLen(pad) <> 1 Then
        Err.Raise 5, "modText", "詰める文字は半角 1 文字にしてください。"
    End If
    SingleBytePad = pad
End Function

Private Function ByteLen(ByVal text As String) As Long
    Dim bytes() As Byte
    If Len(text) = 0 Then Exit Function
    bytes = StrConv(text, vbFromUnicode)
    ByteLen = UBound(bytes) - LBound(bytes) + 1
End Function
