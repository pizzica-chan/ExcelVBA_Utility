Attribute VB_Name = "modDate"
Option Explicit

' 月初月末、会計年度、営業日、日本の祝日。
' このモジュールだけでインポートして使えます。
' 日付を返す関数と IsJapaneseHoliday、JapaneseHolidayName、FiscalYear、AgeYears はワークシート関数としても使えます。
' 日本の祝日は 2020-01-01 から 2099-12-31 です。2020 年と 2021 年の五輪による移動を含みます。
' 春分・秋分は 1980 年から 2099 年向けの近似式です。官報と違う日があれば、祝日セルを holidays に渡してください。
' 2 月 29 日生まれの年齢は、平年では 3 月 1 日を誕生日として数えます。

' 【MonthStart】その月の 1 日。
' 使用例:
'   firstDay = MonthStart(Date)
' 解説: 今日が属する月の 1 日を firstDay に入れる。9 月 22 日なら 9 月 1 日。
' ワークシート関数としても使える。=MonthStart(A2)
Public Function MonthStart(ByVal day As Date) As Date
    MonthStart = DateSerial(Year(day), Month(day), 1)
End Function

' 【MonthEnd】その月の末日。うるう年の 2 月は 29 日。
' 使用例:
'   due = MonthEnd(Range("A2").Value)
' 解説: A2 の日付が属する月の末日を due に入れる。2 月なら、うるう年は 29 日、平年は 28 日。
' ワークシート関数としても使える。=MonthEnd(A2)
Public Function MonthEnd(ByVal day As Date) As Date
    MonthEnd = DateSerial(Year(day), Month(day) + 1, 0)
End Function

' 【WeekStart】週の開始日。既定は月曜始まり。
' 使用例:
'   monday = WeekStart(Date)
'   sundayStart = WeekStart(Date, vbSunday)
' 解説: 1 行目は今日を含む週の月曜日。2 行目は日曜始まりにした、その週の日曜日。
Public Function WeekStart(ByVal day As Date, Optional ByVal firstDay As VbDayOfWeek = vbMonday) As Date
    day = DateOnly(day)
    WeekStart = day - (Weekday(day, firstDay) - 1)
End Function

' 【IsWeekend】土曜または日曜なら True。
' 使用例:
'   If IsWeekend(Range("A2").Value) Then Range("B2").Value = "休日"
' 解説: A2 が土曜または日曜なら、B2 に「休日」と書く。平日のときは B2 を変えない。祝日は見ない。
Public Function IsWeekend(ByVal day As Date) As Boolean
    Dim weekdayNum As Long
    weekdayNum = Weekday(DateOnly(day), vbSunday)
    IsWeekend = (weekdayNum = vbSunday Or weekdayNum = vbSaturday)
End Function

' 【FiscalYear】会計年度。startMonth は開始月で、日本の年度なら 4。
'   2026-03-31 は 2025 年度、2026-04-01 は 2026 年度。
' 使用例:
'   yearNum = FiscalYear(Range("A2").Value, 4)
' 解説: A2 の日付を、4 月始まりの年度で数える。2026 年 3 月 31 日なら 2025、2026 年 4 月 1 日なら 2026。
' ワークシート関数としても使える。=FiscalYear(A2,4)
Public Function FiscalYear(ByVal day As Date, Optional ByVal startMonth As Long = 4) As Long
    If startMonth < 1 Or startMonth > 12 Then Err.Raise 5, "FiscalYear", "開始月は 1 から 12 までです。"
    If Month(day) >= startMonth Then
        FiscalYear = Year(day)
    Else
        FiscalYear = Year(day) - 1
    End If
End Function

' 【FiscalQuarter】開始月からの四半期。4 月開始なら 4-6 月が 1、1-3 月が 4。
' 使用例:
'   q = FiscalQuarter(Date, 4)
' 解説: 今日が、4 月始まりの第何四半期かを q に入れる。4 月から 6 月は 1、1 月から 3 月は 4。
' ワークシート関数としても使える。=FiscalQuarter(A2,4)
Public Function FiscalQuarter(ByVal day As Date, Optional ByVal startMonth As Long = 4) As Long
    Dim offset As Long
    If startMonth < 1 Or startMonth > 12 Then Err.Raise 5, "FiscalQuarter", "開始月は 1 から 12 までです。"
    offset = (Month(day) - startMonth + 12) Mod 12
    FiscalQuarter = offset \ 3 + 1
End Function

' 【YearMonthKey】"yyyy-mm" 形式の年月。集計のキーに使う。
' 使用例:
'   key = YearMonthKey(Range("A2").Value)   ' 2026-09
' 解説: A2 の日付を 2026-09 のような年月の文字列にする。月ごとの集計で、同じ月の行をまとめるキーに使う。
Public Function YearMonthKey(ByVal day As Date) As String
    YearMonthKey = Format$(day, "yyyy-mm")
End Function

' 【AgeYears】満年齢。asOf を省略するか 0 のときは今日。
'   2 月 29 日生まれは、平年では 3 月 1 日を誕生日として数える。
' 使用例:
'   age = AgeYears(Range("C2").Value)
'   age = AgeYears(#2000/4/1#, #2026/3/31#)   ' 25
' 解説: 1 行目は C2 の生年月日から、今日時点の満年齢を求める。2 行目は 2000 年 4 月 1 日生まれを 2026 年 3 月 31 日で見て、誕生日前なので 25 歳。
' ワークシート関数としても使える。=AgeYears(C2)
Public Function AgeYears(ByVal birthDate As Date, Optional ByVal asOf As Date = 0) As Long
    Dim birthdayThisYear As Date
    If asOf = 0 Then asOf = Date
    birthDate = DateOnly(birthDate)
    asOf = DateOnly(asOf)
    AgeYears = DateDiff("yyyy", birthDate, asOf)
    birthdayThisYear = DateSerial(Year(asOf), Month(birthDate), Day(birthDate))
    If birthdayThisYear > asOf Then AgeYears = AgeYears - 1
    If AgeYears < 0 Then AgeYears = 0
End Function

' 【IsJapaneseHoliday】日本の祝日なら True。2020-01-01 から 2099-12-31。
'   振替休日と国民の休日も含む。範囲外の年は False。
' 使用例:
'   If IsJapaneseHoliday(Range("A2").Value) Then Range("B2").Value = "祝日"
' 解説: A2 が 2020 年から 2099 年の日本の祝日なら、B2 に「祝日」と書く。振替休日と国民の休日も含む。平日や範囲外の年は何も書かない。
' ワークシート関数としても使える。=IsJapaneseHoliday(A2)
Public Function IsJapaneseHoliday(ByVal day As Date) As Boolean
    IsJapaneseHoliday = (Len(JapaneseHolidayName(day)) > 0)
End Function

' 【JapaneseHolidayName】祝日名。祝日でなければ空文字。範囲外の年も空文字。
'   2026-09-22 は「国民の休日」、5 月 3 日が日曜の年の振替は「振替休日」。
' 使用例:
'   Range("B2").Value = JapaneseHolidayName(Range("A2").Value)
' 解説: A2 の祝日名を B2 へ書く。元日や成人の日、振替休日など。祝日でなければ B2 は空になる。
' ワークシート関数としても使える。=JapaneseHolidayName(A2)
Public Function JapaneseHolidayName(ByVal day As Date) As String
    Dim yearNum As Long
    Dim key As Long
    Dim map As Object
    day = DateOnly(day)
    yearNum = Year(day)
    If yearNum < 2020 Or yearNum > 2099 Then Exit Function
    Set map = HolidayMap(yearNum)
    key = CLng(day)
    If map.Exists(key) Then JapaneseHolidayName = CStr(map(key))
End Function

' 【JapaneseHolidaysOfYear】1 列目が日付、2 列目が名称の二次元配列。
'   2020 年から 2099 年。春分・秋分は近似式なので、官報と違う年は休日セルを別に持つ。
' 使用例:
'   holidays = JapaneseHolidaysOfYear(2026)
'   Range("A2").Resize(UBound(holidays, 1), 2).Value = holidays
' 解説: 2026 年の祝日を配列で受け取り、A2 から日付と祝日名の 2 列でシートへ書き出す。
Public Function JapaneseHolidaysOfYear(ByVal yearNum As Long) As Variant
    Dim map As Object
    Dim day As Date
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim data() As Variant
    If yearNum < 2020 Or yearNum > 2099 Then
        Err.Raise 5, "JapaneseHolidaysOfYear", "日本の祝日対応は 2020 年から 2099 年までです。"
    End If
    Set map = HolidayMap(yearNum)
    For day = DateSerial(yearNum, 1, 1) To DateSerial(yearNum, 12, 31)
        If map.Exists(CLng(day)) Then rowCount = rowCount + 1
    Next day
    ReDim data(1 To rowCount, 1 To 2)
    For day = DateSerial(yearNum, 1, 1) To DateSerial(yearNum, 12, 31)
        If map.Exists(CLng(day)) Then
            rowIndex = rowIndex + 1
            data(rowIndex, 1) = day
            data(rowIndex, 2) = map(CLng(day))
        End If
    Next day
    JapaneseHolidaysOfYear = data
End Function

' 【WriteJapaneseHolidays】指定年の祝日一覧を、destination の左上から日付・名称の 2 列で書く。
' 使用例:
'   WriteJapaneseHolidays 2026, Worksheets("祝日").Range("A2")
' 解説: 「祝日」シートの A2 から、2026 年の日付と祝日名を 2 列で書く。A1 に見出しがある表の本体として使える。
Public Sub WriteJapaneseHolidays(ByVal yearNum As Long, ByVal destination As Range)
    Dim data As Variant
    data = JapaneseHolidaysOfYear(yearNum)
    destination.Cells(1, 1).Resize(UBound(data, 1), 2).Value = data
End Sub

' 【AddWorkDays】土日を除いた営業日だけ進める。Excel の WORKDAY と同じく、開始日は数えない。
'   useJapaneseHolidays が True のとき祝日も除く。対応は 2020 年から 2099 年。
'   holidays には会社独自の休業日を置く。祝日と併用できる。
' 使用例:
'   due = AddWorkDays(Date, 5, , True)
'   due = AddWorkDays(Range("A2").Value, 3, Worksheets("設定").Range("H2:H20"), True)
' 解説: 1 行目は今日から土日と祝日を除いて 5 営業日後を求める。開始日そのものは数えない。2 行目は A2 から 3 営業日後で、祝日に加えて「設定」の H2:H20 にある会社の休業日も飛ばす。
' ワークシート関数としても使える。=AddWorkDays(A2,5,,TRUE)
Public Function AddWorkDays(ByVal startDate As Date, ByVal days As Long, _
    Optional ByVal holidays As Range, Optional ByVal useJapaneseHolidays As Boolean = False) As Date

    Dim stepDay As Long
    Dim moved As Long
    Dim guard As Long
    Dim cursor As Date
    Dim extra As Object
    startDate = DateOnly(startDate)
    If days = 0 Then
        AddWorkDays = startDate
        Exit Function
    End If
    If useJapaneseHolidays Then EnsureHolidayYear startDate, "AddWorkDays"
    If holidays Is Nothing And Not useJapaneseHolidays Then
        AddWorkDays = DateOnly(Application.WorksheetFunction.WorkDay(startDate, days))
        Exit Function
    End If
    Set extra = HolidayDict(holidays)
    stepDay = IIf(days > 0, 1, -1)
    cursor = startDate
    Do While moved < Abs(days)
        cursor = cursor + stepDay
        If useJapaneseHolidays Then EnsureHolidayYear cursor, "AddWorkDays"
        If IsWorkDayCore(cursor, extra, useJapaneseHolidays) Then moved = moved + 1
        guard = guard + 1
        If guard > 20000 Then Err.Raise 5, "AddWorkDays", "営業日の計算が範囲を超えました。"
    Loop
    AddWorkDays = cursor
End Function

' 【WorkDaysBetween】開始日から終了日までの営業日数。両端を含む。
'   開始が終了より後なら負の数。祝日と追加休業日の扱いは AddWorkDays と同じ。
' 使用例:
'   days = WorkDaysBetween(#2026/5/1#, #2026/5/7#, , True)   ' 2
' 解説: 2026 年 5 月 1 日から 5 月 7 日までを、土日と祝日を除いて数える。両端を含むので、この期間の営業日は 2 日。
Public Function WorkDaysBetween(ByVal startDate As Date, ByVal endDate As Date, _
    Optional ByVal holidays As Range, Optional ByVal useJapaneseHolidays As Boolean = False) As Long

    Dim sign As Long
    Dim day As Date
    Dim extra As Object
    Dim count As Long
    startDate = DateOnly(startDate)
    endDate = DateOnly(endDate)
    sign = 1
    If startDate > endDate Then
        day = startDate
        startDate = endDate
        endDate = day
        sign = -1
    End If
    If holidays Is Nothing And Not useJapaneseHolidays Then
        WorkDaysBetween = Application.WorksheetFunction.NetworkDays(startDate, endDate) * sign
        Exit Function
    End If
    Set extra = HolidayDict(holidays)
    For day = startDate To endDate
        If useJapaneseHolidays Then EnsureHolidayYear day, "WorkDaysBetween"
        If IsWorkDayCore(day, extra, useJapaneseHolidays) Then count = count + 1
    Next day
    WorkDaysBetween = count * sign
End Function

' 【IsWorkDay】土日でなければ True。祝日や追加休業日を渡すとその日は False。
' 使用例:
'   If IsWorkDay(Date, , True) Then MsgBox "今日は営業日です。"
' 解説: 今日が土日でも祝日でもなければ、営業日であるメッセージを出す。会社独自の休業日は見ていない。
Public Function IsWorkDay(ByVal day As Date, Optional ByVal holidays As Range, _
    Optional ByVal useJapaneseHolidays As Boolean = False) As Boolean

    IsWorkDay = IsWorkDayCore(DateOnly(day), HolidayDict(holidays), useJapaneseHolidays)
End Function

' 【Run_WriteHolidaySheet】マクロ一覧用。年を聞いて「祝日2026」のようなシートへ一覧を書く。
'   2020 から 2099。同じ名前のシートがあれば中身を消して書き直す。
' 使用例:
'   実行して 2026 と入力すると、日付と祝日名の一覧ができる。
' 解説: 入力欄に 2026 と入れて OK すると、「祝日2026」シートへその年の祝日一覧を書く。同じ名前のシートがあれば中身を消して書き直す。
Public Sub Run_WriteHolidaySheet()
    Dim yearText As String
    Dim yearNum As Long
    Dim ws As Worksheet
    Dim sheetName As String
    yearText = InputBox("西暦を入力してください（2020-2099）", "祝日一覧", Year(Date))
    If StrPtr(yearText) = 0 Or Len(Trim$(yearText)) = 0 Then Exit Sub
    If Not IsNumeric(yearText) Then
        MsgBox "数値で入力してください。", vbExclamation
        Exit Sub
    End If
    yearNum = CLng(yearText)
    On Error GoTo EH
    sheetName = "祝日" & CStr(yearNum)
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo EH
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    ws.Range("A1").Value = "日付"
    ws.Range("B1").Value = "名称"
    WriteJapaneseHolidays yearNum, ws.Range("A2")
    ws.Columns("A").NumberFormat = "yyyy-mm-dd"
    ws.Columns("A:B").AutoFit
    ws.Range("A1").Select
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "祝日一覧"
End Sub

Private Function IsWorkDayCore(ByVal day As Date, ByVal extra As Object, ByVal useJapaneseHolidays As Boolean) As Boolean
    If IsWeekend(day) Then Exit Function
    If useJapaneseHolidays Then
        If IsJapaneseHoliday(day) Then Exit Function
    End If
    If extra.Exists(CLng(day)) Then Exit Function
    IsWorkDayCore = True
End Function

Private Function HolidayDict(ByVal holidays As Range) As Object
    Dim cell As Range
    Dim value As Variant
    Dim day As Date
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set HolidayDict = result
    If holidays Is Nothing Then Exit Function
    For Each cell In holidays.Cells
        value = cell.Value
        If Not IsEmpty(value) And Not IsError(value) Then
            If IsDate(value) Then
                day = DateOnly(CDate(value))
                result(CLng(day)) = True
            End If
        End If
    Next cell
End Function

Private Sub EnsureHolidayYear(ByVal day As Date, ByVal sourceName As String)
    If Year(day) < 2020 Or Year(day) > 2099 Then
        Err.Raise 5, sourceName, "日本の祝日対応は 2020 年から 2099 年までです。"
    End If
End Sub

Private Function HolidayMap(ByVal yearNum As Long) As Object
    Static cache As Object
    Dim map As Object
    If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
    If cache.Exists(yearNum) Then
        Set HolidayMap = cache(yearNum)
        Exit Function
    End If
    Set map = CreateObject("Scripting.Dictionary")
    AddBaseHolidays map, yearNum
    ApplySubstituteHolidays map, yearNum, False
    ApplyCitizensHolidays map, yearNum
    ApplySubstituteHolidays map, yearNum, True
    cache.Add yearNum, map
    Set HolidayMap = map
End Function

Private Sub AddBaseHolidays(ByVal map As Object, ByVal yearNum As Long)
    AddHoliday map, DateSerial(yearNum, 1, 1), "元日"
    AddHoliday map, NthWeekdayOfMonth(yearNum, 1, 2, vbMonday), "成人の日"
    AddHoliday map, DateSerial(yearNum, 2, 11), "建国記念の日"
    AddHoliday map, DateSerial(yearNum, 2, 23), "天皇誕生日"
    AddHoliday map, DateSerial(yearNum, 3, EquinoxDay(yearNum, 20.8431)), "春分の日"
    AddHoliday map, DateSerial(yearNum, 4, 29), "昭和の日"
    AddHoliday map, DateSerial(yearNum, 5, 3), "憲法記念日"
    AddHoliday map, DateSerial(yearNum, 5, 4), "みどりの日"
    AddHoliday map, DateSerial(yearNum, 5, 5), "こどもの日"
    If yearNum = 2020 Then
        AddHoliday map, DateSerial(yearNum, 7, 23), "海の日"
        AddHoliday map, DateSerial(yearNum, 7, 24), "スポーツの日"
        AddHoliday map, DateSerial(yearNum, 8, 10), "山の日"
    ElseIf yearNum = 2021 Then
        AddHoliday map, DateSerial(yearNum, 7, 22), "海の日"
        AddHoliday map, DateSerial(yearNum, 7, 23), "スポーツの日"
        AddHoliday map, DateSerial(yearNum, 8, 8), "山の日"
    Else
        AddHoliday map, NthWeekdayOfMonth(yearNum, 7, 3, vbMonday), "海の日"
        AddHoliday map, DateSerial(yearNum, 8, 11), "山の日"
        AddHoliday map, NthWeekdayOfMonth(yearNum, 10, 2, vbMonday), "スポーツの日"
    End If
    AddHoliday map, NthWeekdayOfMonth(yearNum, 9, 3, vbMonday), "敬老の日"
    AddHoliday map, DateSerial(yearNum, 9, EquinoxDay(yearNum, 23.2488)), "秋分の日"
    AddHoliday map, DateSerial(yearNum, 11, 3), "文化の日"
    AddHoliday map, DateSerial(yearNum, 11, 23), "勤労感謝の日"
End Sub

Private Sub ApplySubstituteHolidays(ByVal map As Object, ByVal yearNum As Long, ByVal citizensOnly As Boolean)
    Dim day As Date
    Dim cursor As Date
    Dim holidayName As String
    Dim eligible As Boolean
    For day = DateSerial(yearNum, 1, 1) To DateSerial(yearNum, 12, 31)
        If map.Exists(CLng(day)) Then
            holidayName = CStr(map(CLng(day)))
            If citizensOnly Then
                eligible = (holidayName = "国民の休日")
            Else
                eligible = (holidayName <> "振替休日" And holidayName <> "国民の休日")
            End If
            If eligible Then
                If Weekday(day, vbSunday) = vbSunday Then
                    cursor = day + 1
                    Do While map.Exists(CLng(cursor))
                        cursor = cursor + 1
                    Loop
                    If Year(cursor) = yearNum Then
                        If Not map.Exists(CLng(cursor)) Then map.Add CLng(cursor), "振替休日"
                    End If
                End If
            End If
        End If
    Next day
End Sub

Private Sub ApplyCitizensHolidays(ByVal map As Object, ByVal yearNum As Long)
    Dim day As Date
    For day = DateSerial(yearNum, 1, 2) To DateSerial(yearNum, 12, 30)
        If Not map.Exists(CLng(day)) Then
            If map.Exists(CLng(day - 1)) And map.Exists(CLng(day + 1)) Then
                map.Add CLng(day), "国民の休日"
            End If
        End If
    Next day
End Sub

Private Sub AddHoliday(ByVal map As Object, ByVal day As Date, ByVal holidayName As String)
    Dim key As Long
    key = CLng(DateOnly(day))
    If Not map.Exists(key) Then map.Add key, holidayName
End Sub

Private Function NthWeekdayOfMonth(ByVal yearNum As Long, ByVal monthNum As Long, ByVal nth As Long, ByVal dayOfWeek As Long) As Date
    Dim day As Date
    day = DateSerial(yearNum, monthNum, 1)
    Do While Weekday(day, vbSunday) <> dayOfWeek
        day = day + 1
    Loop
    NthWeekdayOfMonth = day + (nth - 1) * 7
End Function

Private Function EquinoxDay(ByVal yearNum As Long, ByVal baseDay As Double) As Long
    Dim delta As Long
    delta = yearNum - 1980
    EquinoxDay = CLng(Int(baseDay + 0.242194 * delta) - (delta \ 4))
End Function

Private Function DateOnly(ByVal value As Date) As Date
    DateOnly = DateSerial(Year(value), Month(value), Day(value))
End Function
