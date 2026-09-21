Attribute VB_Name = "modFile"
Option Explicit

' ファイルとフォルダの確認、選択ダイアログ、一覧。
' このモジュールだけでインポートして使えます。

' 【FileExists】ファイルがあるか。フォルダだけのパスは False。
' 使用例:
'   If FileExists(path) Then Workbooks.Open path
' 解説: path にファイルがあればそのブックを開く。フォルダのパスや、存在しないパスのときは開かない。
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Function
    FileExists = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) > 0)
    If FileExists Then
        FileExists = ((GetAttr(filePath) And vbDirectory) = 0)
    End If
    On Error GoTo 0
End Function

' 【FolderExists】フォルダがあるか。
' 使用例:
'   If Not FolderExists(folder) Then EnsureFolder folder
' 解説: folder がまだ無ければ、そのフォルダを作る。すでにあれば何もしない。
Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    If Len(folderPath) = 0 Then Exit Function
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

' 【EnsureFolder】フォルダを作る。親フォルダが無くても順に作る。既にあれば何もしない。
' 使用例:
'   EnsureFolder CombinePath(ThisFolder(), "出力\2026")
' 解説: マクロのブックと同じ場所に「出力」フォルダと、その中の「2026」フォルダを作る。親が無くてもまとめて作る。
Public Sub EnsureFolder(ByVal folderPath As String)
    Dim parent As String
    folderPath = StripTrailingSlash(folderPath)
    If Len(folderPath) = 0 Then Err.Raise 5, "EnsureFolder", "フォルダが指定されていません。"
    If FolderExists(folderPath) Then Exit Sub
    parent = FolderOf(folderPath)
    If Len(parent) > 0 And StrComp(parent, folderPath, vbTextCompare) <> 0 Then
        If Not IsDriveRoot(parent) Then EnsureFolder parent
    End If
    If Not FolderExists(folderPath) Then MkDir folderPath
End Sub

' 【PickFile】ファイルを 1 つ選ばせる。キャンセルすると空文字。
' 使用例:
'   path = PickFile("マスタを選択", "Excel ブック", "*.xlsx;*.xlsm")
'   If Len(path) = 0 Then Exit Sub
' 解説: 「マスタを選択」というダイアログで xlsx か xlsm を 1 つ選ばせる。選んだパスを path に入れる。キャンセルしたときは空なので、そこで処理をやめる。
Public Function PickFile(Optional ByVal title As String = "ファイルを選択", _
    Optional ByVal filterName As String = "すべてのファイル", _
    Optional ByVal filterPattern As String = "*.*") As String

    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add filterName, filterPattern
        If .Show <> -1 Then Exit Function
        PickFile = .SelectedItems(1)
    End With
End Function

' 【PickFiles】ファイルを複数選ばせる。キャンセルすると空の Collection。
' 使用例:
'   Dim path As Variant
'   For Each path In PickFiles("取り込むファイル", "Excel", "*.xlsx")
'       Debug.Print path
'   Next path
' 解説: Excel ブックを複数選べるダイアログを出し、選んだ各ファイルのパスをイミディエイトへ出す。キャンセルしたときは 1 件も出ない。
Public Function PickFiles(Optional ByVal title As String = "ファイルを選択", _
    Optional ByVal filterName As String = "すべてのファイル", _
    Optional ByVal filterPattern As String = "*.*") As Collection

    Dim dlg As FileDialog
    Dim i As Long
    Set PickFiles = New Collection
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = title
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add filterName, filterPattern
        If .Show <> -1 Then Exit Function
        For i = 1 To .SelectedItems.Count
            PickFiles.Add .SelectedItems(i)
        Next i
    End With
End Function

' 【PickFolder】フォルダを選ばせる。キャンセルすると空文字。
' 使用例:
'   folder = PickFolder("CSV の保存先")
'   If Len(folder) > 0 Then ExportAllSheets ActiveWorkbook, folder
' 解説: 「CSV の保存先」というダイアログでフォルダを選ばせ、選んだ場所へ前面のブックの各シートを CSV で書き出す。キャンセルしたときは何もしない。
Public Function PickFolder(Optional ByVal title As String = "フォルダを選択") As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = title
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        PickFolder = .SelectedItems(1)
    End With
End Function

' 【PickSaveAs】保存先ファイル名を選ばせる。キャンセルすると空文字。
' 使用例:
'   path = PickSaveAs("書き出し先", "売上.csv", "CSV", "*.csv")
' 解説: 初期ファイル名を売上.csv にした保存ダイアログを出す。OK したときの保存先パスが path に入る。キャンセルすると空文字。
Public Function PickSaveAs(Optional ByVal title As String = "保存先を指定", _
    Optional ByVal defaultName As String = "export.csv", _
    Optional ByVal filterName As String = "CSV ファイル", _
    Optional ByVal filterPattern As String = "*.csv") As String

    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    With dlg
        .Title = title
        .InitialFileName = defaultName
        .Filters.Clear
        .Filters.Add filterName, filterPattern
        If .Show <> -1 Then Exit Function
        PickSaveAs = .SelectedItems(1)
    End With
End Function

' 【ListFiles】フォルダ内のファイルパスを Collection で返す。
'   pattern は *.xlsx のような形式。大文字小文字は区別しない。
'   recursive が True のときはサブフォルダも見る。
' 使用例:
'   Dim path As Variant
'   For Each path In ListFiles("C:\inbox", "*.xlsx", True)
'       Set wb = OpenWorkbook(CStr(path), True)
'   Next path
' 解説: C:\inbox とそのサブフォルダにある xlsx を、読み取り専用で順に開く。ファイルが無いフォルダを渡すとエラーになる。
Public Function ListFiles(ByVal folderPath As String, Optional ByVal pattern As String = "*.*", _
    Optional ByVal recursive As Boolean = False) As Collection

    Dim fso As Object
    Set ListFiles = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Err.Raise 76, "ListFiles", "フォルダが見つかりません: " & folderPath
    End If
    CollectFiles fso.GetFolder(folderPath), pattern, recursive, ListFiles
End Function

' 【FileNameOf】パスからファイル名だけを取り出す。
' 使用例:
'   Debug.Print FileNameOf("C:\data\売上.xlsx")   ' 売上.xlsx
' 解説: C:\data\売上.xlsx からフォルダ部分を除き、売上.xlsx だけをイミディエイトへ出す。
Public Function FileNameOf(ByVal filePath As String) As String
    FileNameOf = Mid$(filePath, InStrRev(filePath, "\") + 1)
End Function

' 【BaseNameOf】拡張子を除いたファイル名。
' 使用例:
'   Debug.Print BaseNameOf("C:\data\売上.xlsx")   ' 売上
' 解説: パスから拡張子 .xlsx も除き、売上 だけをイミディエイトへ出す。
Public Function BaseNameOf(ByVal filePath As String) As String
    Dim name As String
    Dim dot As Long
    name = FileNameOf(filePath)
    dot = InStrRev(name, ".")
    If dot > 1 Then
        BaseNameOf = Left$(name, dot - 1)
    Else
        BaseNameOf = name
    End If
End Function

' 【ExtensionOf】拡張子。ドットは含まない。
' 使用例:
'   Debug.Print ExtensionOf("C:\data\売上.xlsx")   ' xlsx
' 解説: 拡張子の xlsx だけを出す。先頭のドットは含まない。
Public Function ExtensionOf(ByVal filePath As String) As String
    Dim name As String
    Dim dot As Long
    name = FileNameOf(filePath)
    dot = InStrRev(name, ".")
    If dot > 0 And dot < Len(name) Then ExtensionOf = Mid$(name, dot + 1)
End Function

' 【FolderOf】ファイルまたはフォルダの親フォルダ。
' 使用例:
'   Debug.Print FolderOf("C:\data\売上.xlsx")   ' C:\data
' 解説: ファイルがある親フォルダ C:\data を出す。ファイル名は含まない。
Public Function FolderOf(ByVal pathText As String) As String
    Dim slash As Long
    pathText = StripTrailingSlash(pathText)
    slash = InStrRev(pathText, "\")
    If slash <= 0 Then Exit Function
    FolderOf = Left$(pathText, slash - 1)
    If Len(FolderOf) = 2 And Mid$(FolderOf, 2, 1) = ":" Then FolderOf = FolderOf & "\"
End Function

' 【CombinePath】フォルダと名前を \ でつなぐ。フォルダ側の末尾 \ はあってもなくてもよい。
' 使用例:
'   path = CombinePath(ThisFolder(), "出力.csv")
' 解説: マクロのブックがあるフォルダと「出力.csv」を \ でつなぎ、保存に使えるフルパスにする。
Public Function CombinePath(ByVal folderPath As String, ByVal name As String) As String
    If Len(folderPath) = 0 Then
        CombinePath = name
    ElseIf Right$(folderPath, 1) = "\" Then
        CombinePath = folderPath & name
    Else
        CombinePath = folderPath & "\" & name
    End If
End Function

' 【SafeFileName】ファイル名に使えない文字を _ にする。空なら "untitled"。
' 使用例:
'   fileName = SafeFileName(ActiveSheet.Name & ".csv")
' 解説: 今のシート名に .csv を付け、ファイル名に使えない \ や : などを _ に直す。シート名が「売上/速報」なら「売上_速報.csv」になる。
Public Function SafeFileName(ByVal fileName As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        fileName = Replace(fileName, CStr(ch), "_")
    Next ch
    fileName = Trim$(fileName)
    If Len(fileName) = 0 Then fileName = "untitled"
    SafeFileName = fileName
End Function

' 【UniqueFilePath】同名ファイルがあるとき、_2, _3 を付けた未使用パスを返す。
' 使用例:
'   path = UniqueFilePath(CombinePath(folder, "売上.csv"))
'   ExportRangeCsv dataBody, path
' 解説: 売上.csv が既にあれば売上_2.csv、それもあれば売上_3.csv のように空いている名前を作り、そのパスへ dataBody を書き出す。
Public Function UniqueFilePath(ByVal filePath As String) As String
    Dim folder As String
    Dim baseName As String
    Dim extension As String
    Dim candidate As String
    Dim i As Long
    If Not FileExists(filePath) Then
        UniqueFilePath = filePath
        Exit Function
    End If
    folder = FolderOf(filePath)
    baseName = BaseNameOf(filePath)
    extension = ExtensionOf(filePath)
    If Len(extension) > 0 Then extension = "." & extension
    For i = 2 To 9999
        candidate = CombinePath(folder, baseName & "_" & CStr(i) & extension)
        If Not FileExists(candidate) Then
            UniqueFilePath = candidate
            Exit Function
        End If
    Next i
    Err.Raise 5, "UniqueFilePath", "空きファイル名を作れません: " & filePath
End Function

' 【DeleteIfExists】ファイルがあれば削除する。無ければ何もしない。
' 使用例:
'   DeleteIfExists CombinePath(TempFolder(), "work.csv")
' 解説: 一時フォルダの work.csv があれば削除する。まだ無ければ何もしない。
Public Sub DeleteIfExists(ByVal filePath As String)
    If FileExists(filePath) Then Kill filePath
End Sub

' 【DesktopFolder】デスクトップのパス。
' 使用例:
'   ExportRangeCsv Selection, CombinePath(DesktopFolder(), "export.csv")
' 解説: 選択範囲を、デスクトップの export.csv として書き出す。
Public Function DesktopFolder() As String
    DesktopFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function

' 【DocumentsFolder】ドキュメント フォルダのパス。
' 使用例:
'   folder = DocumentsFolder()
' 解説: ログイン中のユーザーのドキュメント フォルダのパスを folder に入れる。保存先を固定したいときに使う。
Public Function DocumentsFolder() As String
    DocumentsFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
End Function

' 【TempFolder】一時フォルダのパス。作業ファイルの置き場所に使う。
' 使用例:
'   workFile = CombinePath(TempFolder(), "import.csv")
' 解説: 一時フォルダに import.csv を足したパスを作る。処理の途中で使う作業ファイルの置き場所になる。
Public Function TempFolder() As String
    TempFolder = Environ$("TEMP")
End Function

' 【TimestampedName】名前に日時を付けたファイル名。例: 売上_20260922_030105.xlsx
'   日付と時刻は分けて作る。1 つの書式に月と時を混ぜると、mm が分になることがある。
' 使用例:
'   path = CombinePath(DesktopFolder(), TimestampedName("売上", "xlsx"))
' 解説: デスクトップに、売上_20260922_030105.xlsx のように実行日時を付けたファイル名のパスを作る。同じ名前で上書きしない。
Public Function TimestampedName(ByVal baseName As String, Optional ByVal extension As String = "") As String
    Dim stamp As String
    stamp = Format$(Now, "yyyymmdd") & "_" & Format$(Now, "hhnnss")
    If Len(extension) > 0 Then
        If Left$(extension, 1) <> "." Then extension = "." & extension
    End If
    TimestampedName = SafeFileName(baseName) & "_" & stamp & extension
End Function

Private Sub CollectFiles(ByVal folder As Object, ByVal pattern As String, ByVal recursive As Boolean, ByVal result As Collection)
    Dim file As Object
    Dim subFolder As Object
    For Each file In folder.Files
        If LCase$(file.Name) Like LCase$(pattern) Then result.Add file.Path
    Next file
    If recursive Then
        For Each subFolder In folder.SubFolders
            CollectFiles subFolder, pattern, True, result
        Next subFolder
    End If
End Sub

Private Function StripTrailingSlash(ByVal pathText As String) As String
    pathText = Trim$(pathText)
    Do While Len(pathText) > 3 And Right$(pathText, 1) = "\"
        pathText = Left$(pathText, Len(pathText) - 1)
    Loop
    StripTrailingSlash = pathText
End Function

Private Function IsDriveRoot(ByVal pathText As String) As Boolean
    IsDriveRoot = (Len(pathText) = 3 And Mid$(pathText, 2, 2) = ":\")
End Function
