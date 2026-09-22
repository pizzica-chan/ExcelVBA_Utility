Attribute VB_Name = "modSelfTest"
Option Explicit

' 【SelfTest】全モジュールを入れた状態で動作を確認する。マクロ一覧には出ない。
'   一時ブック上でセル操作、祝日、CSV などを試し、終わるとそのブックは閉じる。
'   失敗するとエラーで止まる。confirm には True を渡す。
' 使用例:
'   SelfTest True
' 解説: イミディエイトウィンドウで実行する。一時ブック上で各機能を試し、問題があればそこでエラーになる。終わると一時ブックと作業ファイルは削除される。引数があるのでマクロ一覧には出ない。
Public Sub SelfTest(ByVal confirm As Boolean)
    Dim wb As Workbook
    Dim prevScreen As Boolean
    Dim prevAlerts As Boolean
    Dim prevEvents As Boolean
    Dim failureNumber As Long
    Dim failureText As String

    If Not confirm Then Err.Raise 5, "SelfTest", "SelfTest True で実行してください。"

    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo EH

    TestPureFunctions
    CleanupTempFolder
    EnsureFolder TempTestFolder()
    TestFiles
    Set wb = Workbooks.Add
    TestSheetsAndBooks wb
    TestRanges wb
    TestLookupTableAndCsv wb
    TestCleanFormatAudit wb
    TestPreflight wb
    wb.Close SaveChanges:=False
    Set wb = Nothing
    CleanupTempFolder

    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    Exit Sub
EH:
    failureNumber = Err.Number
    failureText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    CleanupTempFolder
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    On Error GoTo 0
    Err.Raise failureNumber, "SelfTest", failureText
End Sub

Private Sub TestPureFunctions()
    Dim holidays As Variant
    Dim rowIndex As Long
    Dim foundCitizen As Boolean
    Dim stamp As String

    AssertEqual ColumnLetter(1), "A", "列A"
    AssertEqual ColumnLetter(26), "Z", "列Z"
    AssertEqual ColumnLetter(27), "AA", "列AA"
    AssertEqual ColumnLetter(16384), "XFD", "列XFD"
    AssertEqual ColumnNumber("xfd"), 16384, "列番号 XFD"
    AssertEqual ColumnNumber(3), 3, "列番号 3"

    AssertEqual NormalizeText("a" & ChrW(&H3000) & "  b", True), "a b", "空白の正規化"
    AssertEqual DigitsOnly("TEL０１２-３４"), "01234", "数字だけ"
    AssertEqual ToHalfWidth("Ａ１"), "A1", "半角化"
    AssertEqual ToKatakana("あい"), "アイ", "カタカナ"
    AssertEqual PadLeft("7", 3, "0"), "007", "左埋め"
    AssertEqual Coalesce("", Empty, "x"), "x", "Coalesce"
    AssertEqual Nz(Empty, 0), 0, "Nz"
    AssertEqual CountText("aaa", "aa"), 1, "文字列カウント"
    AssertEqual RegexFirst("abc123xyz", "\d+"), "123", "正規表現"
    AssertEqual RegexReplace("a1b2", "\d", "#"), "a#b#", "正規表現の置換"
    If Asc("あ") < 0 Or Asc("あ") > 255 Then AssertEqual PadLeftB("あ", 5, " "), Space$(3) & "あ", "バイト幅の左埋め"

    AssertEqual CLng(MonthStart(DateSerial(2026, 9, 22))), CLng(DateSerial(2026, 9, 1)), "月初"
    AssertEqual CLng(MonthEnd(DateSerial(2024, 2, 10))), CLng(DateSerial(2024, 2, 29)), "うるう年の月末"
    AssertEqual CLng(MonthEnd(DateSerial(2025, 2, 10))), CLng(DateSerial(2025, 2, 28)), "平年の月末"
    AssertEqual CLng(WeekStart(DateSerial(2026, 9, 22))), CLng(DateSerial(2026, 9, 21)), "週の開始"
    AssertEqual FiscalYear(DateSerial(2026, 3, 31), 4), 2025, "年度の末日"
    AssertEqual FiscalYear(DateSerial(2026, 4, 1), 4), 2026, "年度の開始日"
    AssertEqual FiscalQuarter(DateSerial(2026, 7, 1), 4), 2, "第2四半期"
    AssertEqual FiscalQuarter(DateSerial(2026, 1, 15), 4), 4, "第4四半期"
    AssertEqual AgeYears(DateSerial(2000, 9, 22), DateSerial(2026, 9, 22)), 26, "誕生日当日の年齢"
    AssertEqual AgeYears(DateSerial(2000, 9, 23), DateSerial(2026, 9, 22)), 25, "誕生日前の年齢"

    AssertEqual JapaneseHolidayName(DateSerial(2026, 1, 1)), "元日", "元日"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 1, 12)), "成人の日", "成人の日"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 3, 20)), "春分の日", "2026年の春分"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 5, 6)), "振替休日", "2026年の振替休日"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 9, 21)), "敬老の日", "敬老の日"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 9, 22)), "国民の休日", "国民の休日"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 9, 23)), "秋分の日", "秋分の日"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 10, 12)), "スポーツの日", "スポーツの日"
    AssertEqual JapaneseHolidayName(DateSerial(2024, 2, 12)), "振替休日", "2024年の振替休日"
    AssertEqual JapaneseHolidayName(DateSerial(2024, 3, 20)), "春分の日", "2024年の春分"
    AssertEqual JapaneseHolidayName(DateSerial(2021, 7, 22)), "海の日", "2021年の海の日"
    AssertEqual JapaneseHolidayName(DateSerial(2021, 7, 19)), "", "移動前の海の日"
    AssertEqual JapaneseHolidayName(DateSerial(2021, 8, 8)), "山の日", "2021年の山の日"
    AssertEqual JapaneseHolidayName(DateSerial(2021, 8, 9)), "振替休日", "2021年の山の日の振替"
    AssertEqual JapaneseHolidayName(DateSerial(2026, 6, 1)), "", "平日"
    AssertTrue IsWeekend(DateSerial(2026, 6, 6)), "土曜"
    AssertTrue Not IsWorkDay(DateSerial(2026, 6, 6)), "土曜は営業日ではない"
    AssertTrue IsWorkDay(DateSerial(2026, 6, 1)), "月曜は営業日"

    AssertEqual CLng(AddWorkDays(DateSerial(2026, 5, 1), 1)), CLng(DateSerial(2026, 5, 4)), "週末だけを飛ばす"
    AssertEqual CLng(AddWorkDays(DateSerial(2026, 5, 1), 1, , True)), CLng(DateSerial(2026, 5, 7)), "祝日を飛ばす"
    AssertEqual CLng(AddWorkDays(DateSerial(2026, 5, 7), -1, , True)), CLng(DateSerial(2026, 5, 1)), "営業日を遡る"
    AssertEqual WorkDaysBetween(DateSerial(2026, 5, 1), DateSerial(2026, 5, 7)), 5, "週末だけを除く日数"
    AssertEqual WorkDaysBetween(DateSerial(2026, 5, 1), DateSerial(2026, 5, 7), , True), 2, "祝日を除く日数"
    AssertEqual WorkDaysBetween(DateSerial(2026, 5, 7), DateSerial(2026, 5, 1), , True), -2, "逆順の営業日数"

    holidays = JapaneseHolidaysOfYear(2026)
    AssertEqual UBound(holidays, 1), 18, "2026年の祝日数"
    For rowIndex = 1 To UBound(holidays, 1)
        If CLng(holidays(rowIndex, 1)) = CLng(DateSerial(2026, 9, 22)) Then
            AssertEqual CStr(holidays(rowIndex, 2)), "国民の休日", "一覧の国民の休日"
            foundCitizen = True
        End If
    Next rowIndex
    AssertTrue foundCitizen, "2026-09-22 が一覧にある"

    AssertEqual SafeFileName("a/b:c"), "a_b_c", "ファイル名の禁止文字"
    AssertEqual CombinePath("C:\temp", "a.txt"), "C:\temp\a.txt", "パス結合"
    AssertEqual CombinePath("C:\temp\", "a.txt"), "C:\temp\a.txt", "区切り済みパス結合"
    AssertEqual BaseNameOf("C:\a\b.xlsx"), "b", "拡張子を除く名前"
    AssertEqual ExtensionOf("C:\a\b.xlsx"), "xlsx", "拡張子"
    AssertEqual FolderOf("C:\a\b.xlsx"), "C:\a", "親フォルダ"
    AssertEqual FolderOf("C:\a"), "C:\", "ドライブ直下の親"
    stamp = TimestampedName("売上", "xlsx")
    AssertTrue Left$(stamp, 3) = "売上_", "タイムスタンプ名の接頭辞"
    AssertTrue InStr(stamp, Format$(Date, "yyyymmdd")) > 0, "タイムスタンプ名の日付"
    AssertTrue Right$(stamp, 5) = ".xlsx", "タイムスタンプ名の拡張子"
    AssertTrue Len(DesktopFolder()) > 0, "デスクトップ"
    AssertTrue Not IsWorkbookOpen("this-file-is-not-open.xlsx"), "開いていないブック"
End Sub

Private Sub TestFiles()
    Dim folder As String
    Dim filePath As String
    Dim nested As String
    Dim listed As Collection
    Dim handle As Integer

    folder = TempTestFolder()
    nested = CombinePath(folder, "nested\child")
    EnsureFolder nested
    AssertTrue FolderExists(nested), "フォルダ作成"

    filePath = CombinePath(folder, "sample.txt")
    handle = FreeFile
    Open filePath For Output As #handle
    Print #handle, "x"
    Close #handle
    AssertTrue FileExists(filePath), "ファイルがある"
    AssertTrue Not FileExists(CombinePath(folder, "missing.txt")), "ファイルが無い"
    AssertEqual UniqueFilePath(filePath), CombinePath(folder, "sample_2.txt"), "重ならないパス"
    Set listed = ListFiles(folder, "*.txt", False)
    AssertEqual listed.Count, 1, "ファイル一覧"
    DeleteIfExists filePath
    AssertTrue Not FileExists(filePath), "ファイル削除"
End Sub

Private Sub TestSheetsAndBooks(ByVal wb As Workbook)
    Dim extra As Workbook
    Dim savedPath As String
    Dim opened As Workbook

    wb.Worksheets(1).Name = "m_sort"
    GetOrCreateSheet "c_sort", wb
    GetOrCreateSheet "a_sort", wb
    GetOrCreateSheet "b_sort", wb
    AssertTrue SheetExists("a_sort", wb), "シートがある"
    AssertEqual GetOrCreateSheet("a_sort", wb).Name, "a_sort", "既存シートを返す"
    AssertEqual SafeSheetName("ab:cd/ef"), "ab_cd_ef", "シート名の禁止文字"
    AssertTrue UniqueSheetName(wb, "a_sort") <> "a_sort", "重ならないシート名"
    SortSheetsByName wb
    AssertEqual wb.Worksheets(1).Name, "a_sort", "シート順 1"
    AssertEqual wb.Worksheets(2).Name, "b_sort", "シート順 2"
    AssertEqual wb.Worksheets(3).Name, "c_sort", "シート順 3"
    AssertEqual wb.Worksheets(4).Name, "m_sort", "シート順 4"
    AssertTrue DeleteSheet("b_sort", wb), "シート削除"
    AssertTrue Not SheetExists("b_sort", wb), "削除済み"
    AssertEqual GetOrCreateSheet("x:y/z", wb).Name, "x_y_z", "不正なシート名を直して作る"

    SetVeryHidden wb.Worksheets("m_sort")
    AssertEqual wb.Worksheets("m_sort").Visible, xlSheetVeryHidden, "非常に非表示"
    ShowAllSheets wb
    AssertEqual wb.Worksheets("m_sort").Visible, xlSheetVisible, "再表示"

    savedPath = CombinePath(TempTestFolder(), "open_test.xlsx")
    DeleteIfExists savedPath
    Set extra = Workbooks.Add
    extra.Worksheets(1).Range("A1").Value = "ping"
    extra.SaveAs Filename:=savedPath, FileFormat:=xlOpenXMLWorkbook
    extra.Close SaveChanges:=False
    Set opened = OpenWorkbook(savedPath)
    AssertEqual opened.Worksheets(1).Range("A1").Value, "ping", "ブックを開く"
    AssertTrue IsWorkbookOpen(savedPath), "開いたブックを見つける"
    opened.Close SaveChanges:=False
    DeleteIfExists savedPath

    Set extra = CopySheetToNewWorkbook(wb.Worksheets("a_sort"))
    AssertTrue extra.Worksheets.Count >= 1, "シートを新しいブックへ"
    extra.Close SaveChanges:=False

    ResetViewToA1 wb
    ProtectAllSheets wb, "pw"
    AssertTrue wb.Worksheets("a_sort").ProtectContents, "シート保護"
    UnprotectAllSheets wb, "pw"
    AssertTrue Not wb.Worksheets("a_sort").ProtectContents, "保護解除"
End Sub

Private Sub TestRanges(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim found As Collection

    Set ws = GetOrCreateSheet("rng", wb)
    ws.Range("A1").Value = "a"
    ws.Range("A2").Value = ""
    ws.Range("A3").Value = "b"
    DeleteBlankRows ws.Range("A1:A3")
    AssertEqual ws.Range("A1").Value, "a", "空白行削除後の1行目"
    AssertEqual ws.Range("A2").Value, "b", "空白行削除後の2行目"
    AssertEqual LastRow(ws, "A"), 2, "列の最終行"

    Set ws = GetOrCreateSheet("fill", wb)
    ws.Range("A1").Value = "x"
    ws.Range("A2").Value = ""
    ws.Range("A3").Value = ""
    FillDownBlanks ws.Range("A1:A3")
    AssertEqual ws.Range("A2").Value, "x", "上の値で埋める"
    AssertEqual ws.Range("A3").Value, "x", "連続した空欄を埋める"

    Set ws = GetOrCreateSheet("merge", wb)
    ws.Range("A1").Value = "m"
    ws.Range("A1:A2").Merge
    UnmergeAndFill ws.Range("A1:A2")
    AssertTrue Not ws.Range("A1").MergeCells, "結合解除"
    AssertEqual ws.Range("A1").Value, "m", "結合解除後の上セル"
    AssertEqual ws.Range("A2").Value, "m", "結合解除後の下セル"

    Set ws = GetOrCreateSheet("calc", wb)
    ws.Range("A1").Formula = "=1+1"
    CopyValues ws.Range("A1"), ws.Range("B1")
    AssertEqual ws.Range("B1").Value, 2, "値の転記"
    AssertTrue Not ws.Range("B1").HasFormula, "転記先は数式ではない"
    PasteValuesInPlace ws.Range("A1")
    AssertTrue Not ws.Range("A1").HasFormula, "値貼り"
    AssertEqual ws.Range("A1").Value, 2, "値貼り後の値"
    ws.Range("C1").Value = "cat"
    ws.Range("C2").Value = "dog"
    ws.Range("C3").Value = "cat"
    Set found = FindAll(ws.Range("C1:C3"), "cat")
    AssertEqual found.Count, 2, "すべて検索"
    HighlightDuplicates ws.Range("C1:C3")
    AssertTrue ws.Range("C1:C3").FormatConditions.Count >= 1, "重複の色"

    Set ws = GetOrCreateSheet("filter", wb)
    ws.Range("A1").Value = "Kind"
    ws.Range("A2").Value = "x"
    ws.Range("A3").Value = "y"
    FilterByHeader ws.Range("A1"), "x"
    ShowAllData ws
    AssertTrue Not ws.FilterMode, "フィルタ解除"
    AssertEqual DataRange(ws).Address(False, False), "A1:A3", "データ範囲"
End Sub

Private Sub TestLookupTableAndCsv(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim dict As Object
    Dim table As ListObject
    Dim data As Variant
    Dim matrix As Variant
    Dim filePath As String

    Set ws = GetOrCreateSheet("grid", wb)
    ws.Range("A1").Value = 1
    ws.Range("B1").Value = 2
    ws.Range("A2").Value = 3
    ws.Range("B2").Value = 4
    matrix = ToMatrix(ws.Range("A1"))
    AssertEqual MatrixRows(matrix), 1, "1セルの行数"
    AssertEqual MatrixCols(matrix), 1, "1セルの列数"
    WriteMatrix TransposeMatrix(ToMatrix(ws.Range("A1:B2"))), ws.Range("D1")
    AssertEqual ws.Range("D1").Value, 1, "転置 1"
    AssertEqual ws.Range("E1").Value, 3, "転置 2"
    AssertEqual ws.Range("D2").Value, 2, "転置 3"
    AssertEqual ws.Range("E2").Value, 4, "転置 4"

    ws.Range("A10").Value = "Ab"
    ws.Range("B10").Value = 10
    ws.Range("A11").Value = "k"
    ws.Range("B11").Value = 20
    Set dict = LoadDictionary(ws.Range("A10:A11"), ws.Range("B10:B11"), 1)
    AssertEqual DictItem(dict, "ab", ""), 10, "大文字小文字を無視して引く"
    AssertEqual DictItem(dict, "missing", "no"), "no", "見つからないキー"
    ws.Range("C10").Value = "k"
    ws.Range("C11").Value = "missing"
    WriteLookup ws.Range("C10:C11"), dict, ws.Range("D10"), "no"
    AssertEqual ws.Range("D10").Value, 20, "辞書から転記"
    AssertEqual ws.Range("D11").Value, "no", "見つからないときの値"

    Set ws = GetOrCreateSheet("items", wb)
    ws.Range("A1").Value = "Name"
    ws.Range("B1").Value = "Qty"
    ws.Range("A2").Value = "a"
    ws.Range("B2").Value = 1
    Set table = EnsureTable(ws.Range("A1:B2"), "Sales")
    AssertEqual TableColumnNumber(table, "Qty"), 2, "テーブル列番号"
    AppendTableRow table, Array("b", 2)
    data = ReadTable(table, True)
    AssertEqual UBound(data, 1), 3, "テーブルの行数"
    AssertEqual data(3, 1), "b", "追加した行"
    AssertEqual data(3, 2), 2, "追加した数量"
    ClearTableRows table
    AssertEqual table.ListRows.Count, 0, "テーブルのデータを消す"

    Set ws = GetOrCreateSheet("csv", wb)
    ws.Range("A1").Value = "あい"
    ws.Range("B1").Value = "a""b,c"
    ws.Range("A2").Value = DateSerial(2026, 9, 22)
    ws.Range("B2").Value = 1234.5
    ws.Range("A3").Value = "line1" & vbLf & "line2"
    filePath = CombinePath(TempTestFolder(), "roundtrip.csv")
    ExportRangeCsv ws.Range("A1:B3"), filePath, "UTF-8"
    ImportCsv filePath, ws.Range("D1"), "UTF-8"
    AssertEqual ws.Range("D1").Value, "あい", "CSVの日本語"
    AssertEqual ws.Range("E1").Value, "a""b,c", "CSVの引用符"
    AssertEqual CLng(ws.Range("D2").Value), CLng(DateSerial(2026, 9, 22)), "CSVの日付"
    AssertClose CDbl(ws.Range("E2").Value), 1234.5, "CSVの数値"
    AssertEqual ws.Range("D3").Value, "line1" & vbLf & "line2", "CSVの改行"
End Sub

Private Sub TestCleanFormatAudit(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim hits As Variant
    Dim removed As Long
    Dim written As Long

    Set ws = GetOrCreateSheet("clean", wb)
    ws.Range("A1").Value = "  あい" & ChrW(&H3000)
    TrimCells ws.Range("A1")
    AssertEqual ws.Range("A1").Value, "あい", "セルの空白除去"
    ws.Range("A2:A3").NumberFormat = "@"
    ws.Range("A2").Value = "123"
    ws.Range("A3").Value = "012"
    FixNumbersStoredAsText ws.Range("A2:A3")
    AssertEqual CDbl(ws.Range("A2").Value), 123, "数値テキストの変換"
    AssertTrue VarType(ws.Range("A2").Value) <> vbString, "数値として保存"
    AssertEqual CStr(ws.Range("A3").Value), "012", "先頭ゼロは残す"
    ws.Range("C1").Value = "foo bar foo"
    ReplaceText ws.Range("C1"), "foo", "baz"
    AssertEqual ws.Range("C1").Value, "baz bar baz", "文字列置換"
    ws.Range("B1").Value = "Header"
    ws.Range("B2").Value = "a"
    ws.Range("B3").Value = "a"
    ws.Range("B4").Value = "b"
    removed = DeduplicateRows(ws.Range("B1:B4"), xlYes)
    AssertEqual removed, 1, "重複行の削除数"
    AssertEqual ws.Range("B2").Value, "a", "重複削除後の値"
    AssertEqual ws.Range("B3").Value, "b", "残った行"
    ws.Hyperlinks.Add Anchor:=ws.Range("D1"), Address:="https://example.com", TextToDisplay:="ex"
    RemoveHyperlinks ws.Range("D1")
    AssertEqual ws.Range("D1").Hyperlinks.Count, 0, "ハイパーリンク削除"

    StyleHeaderRow ws.Range("A5:B5")
    AssertTrue ws.Range("A5").Font.Bold, "見出しを太字"
    ws.Range("A6").Value = DateSerial(2026, 9, 22)
    FormatAsDate ws.Range("A6")
    AssertTrue ws.Range("A6").NumberFormat <> "General", "日付書式"
    ApplyGridBorders ws.Range("A5:B6")
    StripeRows ws.Range("A5:B8")
    AssertTrue ws.Range("A5:B8").FormatConditions.Count >= 1, "縞模様"
    AutoFitLimited ws.Range("A5:B6")

    written = WriteSheetIndex(wb, ws.Range("F1"))
    AssertTrue written >= 2, "シート一覧"
    AssertEqual ws.Range("F1").Value, "シート名", "シート一覧の見出し"
    ws.Range("H1").Formula = "=1+1"
    written = WriteFormulaList(ws.Range("H1"), ws.Range("H3"))
    AssertTrue InStr(CStr(ws.Range("I4").Value), "1+1") > 0, "数式一覧"
    written = WriteDefinedNames(wb, ws.Range("F20"))
    AssertEqual ws.Range("F20").Value, "名前", "名前定義の見出し"

    ws.Range("A20").Value = "xxCellWordXYZ99"
    ws.Range("Z1").Value = "CellWordXYZ99"
    hits = FindForbiddenWords(ws.Range("Z1"), wb)
    AssertTrue HasForbiddenHit(hits, "セル", "A20", "CellWordXYZ99"), "セルの禁止ワード"
    AssertTrue Not HasForbiddenHit(hits, "セル", "Z1", "CellWordXYZ99"), "一覧セルは除外"
    hits = FindForbiddenWords(Array("cellwordxyz99"), wb)
    AssertTrue HasForbiddenHit(hits, "セル", "A20", "cellwordxyz99"), "大文字小文字を無視"
    ws.Range("A23").Formula = "=1+1+N(""FormulaWordXYZ99"")"
    hits = FindForbiddenWords(Array("FormulaWordXYZ99"), wb)
    AssertTrue HasForbiddenHit(hits, "数式", "A23", "FormulaWordXYZ99"), "数式内の禁止ワード"
    ws.Range("A24").AddComment "NoteWordXYZ99"
    hits = FindForbiddenWords(Array("NoteWordXYZ99"), wb)
    AssertTrue HasForbiddenHit(hits, "メモ", "A24", "NoteWordXYZ99"), "メモの禁止ワード"
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 120, 30)
    shp.TextFrame.Characters.Text = "ShapeWordXYZ99"
    hits = FindForbiddenWords(Array("ShapeWordXYZ99"), wb)
    AssertTrue HasForbiddenHit(hits, "図形", shp.Name, "ShapeWordXYZ99"), "図形の禁止ワード"
    written = WriteForbiddenWords(Array("FormulaWordXYZ99"), ws.Range("AA1"), wb)
    AssertEqual ws.Range("AA1").Value, "種別", "禁止ワード結果の見出し"
End Sub

Private Sub TestPreflight(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim secret As Worksheet
    Dim shp As Shape
    Dim hits As Variant
    Dim authorSet As Boolean
    Dim headerSet As Boolean

    Set ws = GetOrCreateSheet("pii", wb)
    ws.Range("A1:A6").NumberFormat = "@"
    ws.Range("A1").Value = "連絡は test@example.com まで"
    ws.Range("A2").Value = "090-1234-5678"
    ws.Range("A3").Value = "〒100-0001"
    ws.Range("A4").Value = "4111111111111111"
    ws.Range("A5").Value = "123456789018"
    ws.Range("A6").Value = "０９０－１２３４－５６７８"
    ws.Range("B1").Value = "2026-09-22"
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = "123456789010"
    ws.Range("B3").NumberFormat = "@"
    ws.Range("B3").Value = "4111111111111112"
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 160, 30)
    shp.TextFrame.Characters.Text = "shape@example.com"
    hits = FindPersonalData(wb)
    AssertTrue HasInspectHit(hits, "メール", "pii", "セル A1", "t***@example.com"), "メール"
    AssertTrue HasInspectHit(hits, "電話", "pii", "セル A2", "09*-****-**78"), "電話"
    AssertTrue HasInspectHit(hits, "郵便番号", "pii", "セル A3", "〒100-0001"), "郵便番号"
    AssertTrue HasInspectHit(hits, "カード番号", "pii", "セル A4", "************1111"), "カード番号"
    AssertTrue HasInspectHit(hits, "個人番号", "pii", "セル A5", "**********18"), "個人番号"
    AssertTrue HasInspectHit(hits, "電話", "pii", "セル A6", "09*-****-**78"), "全角の電話"
    AssertTrue HasInspectHit(hits, "メール", "pii", "図形 " & shp.Name, "s***@example.com"), "図形のメール"
    AssertTrue Not HasInspectHit(hits, "*", "pii", "セル B1", "*"), "日付は該当にしない"
    AssertTrue Not HasInspectHit(hits, "*", "pii", "セル B2", "*"), "検査数字の違う個人番号"
    AssertTrue Not HasInspectHit(hits, "*", "pii", "セル B3", "*"), "検査数字の違うカード番号"

    Set ws = GetOrCreateSheet("inspect", wb)
    ws.Range("A40").Value = "HiddenRowXYZ"
    ws.Rows(40).Hidden = True
    ws.Range("C5").Value = "WhiteXYZ"
    ws.Range("C5").Font.Color = RGB(255, 255, 255)
    ws.Range("D5").Value = "VisibleWhiteXYZ"
    ws.Range("D5").Font.Color = RGB(255, 255, 255)
    ws.Range("D5").Interior.Color = RGB(0, 32, 96)
    ws.Range("C6").Value = "TinyXYZ"
    ws.Range("C6").Font.Size = 1
    wb.Names.Add Name:="SecretNameXYZ", RefersTo:="='" & ws.Name & "'!$A$40", Visible:=False
    Set secret = GetOrCreateSheet("secretbox", wb)
    secret.Range("A1").Value = "secret"
    secret.Visible = xlSheetVeryHidden
    hits = FindHiddenContent(wb)
    AssertTrue HasInspectHit(hits, "非表示行", "inspect", "40", "HiddenRowXYZ"), "非表示行"
    AssertTrue HasInspectHit(hits, "白文字", "inspect", "C5", "WhiteXYZ"), "白文字"
    AssertTrue Not HasInspectHit(hits, "白文字", "inspect", "D5", "*"), "濃い背景の白文字"
    AssertTrue HasInspectHit(hits, "極小文字", "inspect", "C6", "TinyXYZ"), "極小文字"
    AssertTrue HasInspectHit(hits, "非常に非表示", "secretbox", "*", "*"), "非常に非表示"
    AssertTrue HasInspectHit(hits, "非表示の名前", "inspect", "SecretNameXYZ", "*"), "非表示の名前"
    AssertTrue HasInspectHit(hits, "参照先", "inspect", "SecretNameXYZ", "非表示行を参照"), "名前の参照先"

    authorSet = False
    headerSet = False
    On Error Resume Next
    wb.BuiltinDocumentProperties("Author").Value = "AuthorXYZ99"
    authorSet = (Err.Number = 0)
    Err.Clear
    Application.PrintCommunication = False
    Err.Clear
    ws.PageSetup.LeftHeader = "HeaderXYZ99"
    headerSet = (Err.Number = 0)
    Err.Clear
    Application.PrintCommunication = True
    On Error GoTo 0
    hits = FindWorkbookProfile(wb)
    AssertTrue HitContentHas(hits, wb.Name), "ファイル名"
    If authorSet Then AssertTrue HitContentHas(hits, "AuthorXYZ99"), "作成者"
    If headerSet Then AssertTrue HitContentHas(hits, "HeaderXYZ99"), "ヘッダー"
End Sub

Private Function HasInspectHit(ByVal hits As Variant, ByVal kind As String, ByVal sheetName As String, _
    ByVal place As String, ByVal content As String) As Boolean

    Dim rowIndex As Long
    If IsEmpty(hits) Then Exit Function
    For rowIndex = 1 To UBound(hits, 1)
        If InspectFieldOk(kind, hits(rowIndex, 1)) And InspectFieldOk(sheetName, hits(rowIndex, 2)) _
            And InspectFieldOk(place, hits(rowIndex, 3)) And InspectFieldOk(content, hits(rowIndex, 4)) Then
            HasInspectHit = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function InspectFieldOk(ByVal expected As String, ByVal actual As Variant) As Boolean
    Dim actualText As String
    If expected = "*" Then
        InspectFieldOk = True
        Exit Function
    End If
    If IsEmpty(actual) Then
        actualText = ""
    Else
        actualText = CStr(actual)
    End If
    InspectFieldOk = (expected = actualText)
End Function

Private Function HitContentHas(ByVal hits As Variant, ByVal fragment As String) As Boolean
    Dim rowIndex As Long
    Dim text As String
    If IsEmpty(hits) Then Exit Function
    For rowIndex = 1 To UBound(hits, 1)
        If IsEmpty(hits(rowIndex, 4)) Then
            text = ""
        Else
            text = CStr(hits(rowIndex, 4))
        End If
        If InStr(1, text, fragment, vbTextCompare) > 0 Then
            HitContentHas = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function HasForbiddenHit(ByVal hits As Variant, ByVal kind As String, ByVal place As String, ByVal word As String) As Boolean
    Dim rowIndex As Long
    If IsEmpty(hits) Then Exit Function
    For rowIndex = 1 To UBound(hits, 1)
        If CStr(hits(rowIndex, 1)) = kind And CStr(hits(rowIndex, 3)) = place And CStr(hits(rowIndex, 4)) = word Then
            HasForbiddenHit = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function TempTestFolder() As String
    TempTestFolder = CombinePath(TempFolder(), "VBA_Lib_SelfTest")
End Function

Private Sub CleanupTempFolder()
    Dim folder As String
    Dim listed As Collection
    Dim path As Variant
    Dim nested As String
    folder = TempTestFolder()
    If Not FolderExists(folder) Then Exit Sub
    Set listed = ListFiles(folder, "*.*", True)
    For Each path In listed
        DeleteIfExists CStr(path)
    Next path
    nested = CombinePath(folder, "nested\child")
    On Error Resume Next
    If FolderExists(nested) Then RmDir nested
    nested = CombinePath(folder, "nested")
    If FolderExists(nested) Then RmDir nested
    RmDir folder
    On Error GoTo 0
End Sub

Private Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal message As String)
    Dim actualText As String
    Dim expectedText As String
    actualText = VariantText(actual)
    expectedText = VariantText(expected)
    If actualText <> expectedText Then
        Err.Raise vbObjectError + 1, "SelfTest", message & " / 期待: [" & expectedText & "] / 実際: [" & actualText & "]"
    End If
End Sub

Private Sub AssertClose(ByVal actual As Double, ByVal expected As Double, ByVal message As String)
    If Abs(actual - expected) > 0.0000001 Then
        Err.Raise vbObjectError + 1, "SelfTest", message & " / 期待: " & CStr(expected) & " / 実際: " & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 1, "SelfTest", message
End Sub

Private Function VariantText(ByVal value As Variant) As String
    If IsEmpty(value) Then
        VariantText = "<Empty>"
    ElseIf IsNull(value) Then
        VariantText = "<Null>"
    ElseIf IsError(value) Then
        VariantText = "<Error>"
    Else
        VariantText = CStr(value)
    End If
End Function
