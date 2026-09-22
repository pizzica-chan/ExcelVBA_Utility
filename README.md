# ExcelVBA_Utility

Excel で繰り返し使う処理を、標準モジュールに分けて置いています。ライブラリとしての名前は VBA_Lib で、取り込むと `dist\VBA_Lib.xlsm` ができます。必要なファイルだけをブックへ取り込めます。参照設定は不要です。

`Run_` で始まるものだけがマクロ一覧に出ます。それ以外は、いつものマクロから関数として呼びます。

## 取り込み

推奨はスクリプトです。ソースは UTF-8 なので、VBE の「ファイルのインポート」だと日本語コメントが文字化けします。

1. Excel の「ファイル」→「オプション」→「トラスト センター」→「トラスト センターの設定」→「マクロの設定」で、「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオンにする。
2. PowerShell で次を実行する。

```powershell
powershell -ExecutionPolicy Bypass -File tools\Import-VbaLibrary.ps1
```

`dist\VBA_Lib.xlsm` ができます。取り込みのあと、日付・CSV・シート操作などの自己テストも実行します。

手動で足すときは、`.bas` をエディタで開き、中身を標準モジュールへ貼り付けてください。クリップボード経由なら日本語が崩れません。

## 呼び出し例

```vba
Debug.Print LastRow(ActiveSheet, "B")
Debug.Print JapaneseHolidayName(Date)
GetOrCreateSheet "集計"
ExportRangeCsv Selection, "C:\temp\out.csv"
```

ワークシートの数式にも使えます。

```excel
=FiscalYear(A1, 4)
=JapaneseHolidayName(A1)
=AgeYears(A1)
=IsJapaneseHoliday(A1)
```

## マクロ一覧に出るもの

選択範囲やアクティブシートに対して動きます。

| マクロ | 内容 |
| --- | --- |
| `Run_DeleteBlankRows` | 選択範囲内が空の行を、行ごと削除する |
| `Run_ConvertSelectionToValues` | 選択範囲の数式を値にする |
| `Run_HighlightDuplicates` | 重複値に色を付ける |
| `Run_UnmergeAndFill` | 結合を解除し、値を各セルへ埋める |
| `Run_FillDownBlanks` | 空欄を上の値で埋める |
| `Run_ClearFilter` | 絞り込みを解除する |
| `Run_SortSheetNames` | シート名で並べ替える |
| `Run_ResetViewToA1` | 表示中のシートを A1 に戻す |
| `Run_ShowAllSheets` | 非表示シートをすべて表示する |
| `Run_ProtectAllSheets` | 全シートを保護する |
| `Run_UnprotectAllSheets` | 全シートの保護を解除する |
| `Run_SaveAllWorkbooks` | 保存済みのブックを保存する |
| `Run_RefreshWorkbook` | クエリとピボットを更新する |
| `Run_AutoFitSelection` | 列幅を内容に合わせ、幅の上下限を付ける |
| `Run_FreezeTopRow` | 先頭行でウィンドウ枠を固定する |
| `Run_PrintSetupA4` | A4・1 ページ幅の印刷設定にする |
| `Run_ExportSelectionToCsv` | 選択範囲を UTF-8 の CSV にする |
| `Run_ImportCsvToSelection` | CSV を選択セルへ読み込む |
| `Run_TrimSelection` | 選択範囲の前後の空白を除く |
| `Run_FixNumericText` | 数値テキストを数値にする |
| `Run_RemoveDuplicates` | 選択範囲の重複行を削除する |
| `Run_RemoveHyperlinks` | ハイパーリンクを外す |
| `Run_WriteHolidaySheet` | 指定年の祝日一覧シートを作る |
| `Run_WriteAuditSheets` | シート、名前、リンク、数式の点検シートを作る |
| `Run_FindForbiddenWords` | 禁止ワードがセルやオブジェクトに含まれるか調べる |
| `Run_FindPersonalData` | メール、電話、郵便番号、カード番号、個人番号に見える文字列を探す |
| `Run_FindHiddenContent` | 非表示のシート・行・列、白文字、極小文字、非表示の名前を探す |
| `Run_WriteWorkbookProfile` | 作成者、ヘッダー、外部リンク、接続を一覧する |

## モジュール

### modRange（セル範囲）

- `LastRow` / `LastColumn` … 最終行・最終列。データが無いときは 0
- `DataRange` … A1 から最終セルまでの範囲
- `CurrentDataRegion` … 連続したデータ領域
- `ColumnLetter` / `ColumnNumber` … 列記号と列番号の変換
- `CopyValues` / `PasteValuesInPlace` … 値だけコピー、数式を値に変換
- `FindAll` … 一致セルをすべて返す
- `DeleteBlankRows` … 範囲内が空の行を、シートの行ごと削除する
- `FillDownBlanks` … 空欄を上の値で埋める
- `UnmergeAndFill` … 結合解除して値を埋める
- `HighlightDuplicates` … 重複に条件付き書式を付ける
- `FilterByHeader` / `ShowAllData` … 見出し列で絞る、絞り込みを解除する

`DeleteBlankRows` は対象の行全体を消します。範囲の外にある同じ行の値も消えます。

### modSheet（シート）

- `SheetExists` / `GetSheet` / `GetOrCreateSheet` / `DeleteSheet`
- `RenameSheet` / `SafeSheetName` / `UniqueSheetName` / `SheetNames`
- `SortSheetsByName`
- `ProtectAllSheets` / `UnprotectAllSheets`
- `ResetViewToA1` / `ShowAllSheets` / `SetVeryHidden`
- `CopySheetToNewWorkbook`

`ProtectAllSheets` は `UserInterfaceOnly` を付けます。開いている間はマクロから編集できます。ブックを開き直すとこの指定は消えるので、その後もマクロから書きたいときは `Workbook_Open` で同じ保護をかけ直してください。`Run_ProtectAllSheets` と `Run_UnprotectAllSheets` の入力欄はパスワードをそのまま表示します。画面に残したくないときは、それぞれの関数をコードから呼んでください。

### modBook（ブック）

- `FindWorkbook` / `IsWorkbookOpen` / `OpenWorkbook`
- `ThisFolder` / `ActiveFolder` … 未保存のブックでは空文字
- `SaveAllWorkbooks` … 未保存・読み取り専用・アドインは保存しない
- `CloseOtherWorkbooks` … このマクロが入っているブック以外を閉じる。保存するなら引数を True にする
- `RefreshWorkbook` … クエリとピボットを更新し、非同期クエリの完了を待つ
- `ExternalLinks`

### modFile（ファイル）

- `FileExists` / `FolderExists` / `EnsureFolder`
- `PickFile` / `PickFiles` / `PickFolder` / `PickSaveAs`
- `ListFiles` … `*.xlsx` のようなパターン。サブフォルダも指定できる
- `FileNameOf` / `BaseNameOf` / `ExtensionOf` / `FolderOf` / `CombinePath`
- `SafeFileName` / `UniqueFilePath` / `DeleteIfExists`
- `DesktopFolder` / `DocumentsFolder` / `TempFolder`
- `TimestampedName` … `売上_20260922_030105.xlsx` のような名前

### modText（文字列）

- `Nz` / `IsBlank` / `Coalesce`
- `NormalizeText` / `CollapseSpaces`
- `DigitsOnly` … 全角数字も拾う
- `ToHalfWidth` / `ToFullWidth` / `ToHiragana` / `ToKatakana`
- `PadLeft` / `PadRight` / `PadLeftB` / `PadRightB`
- `CountText`
- `RegexFirst` / `RegexMatch` / `RegexReplace`

`PadLeftB` のバイト数は、日本語 Windows の Shift_JIS 換算です。詰める文字は半角 1 文字にしてください。

### modDate（日付・営業日・祝日）

- `MonthStart` / `MonthEnd` / `WeekStart` / `IsWeekend`
- `FiscalYear` / `FiscalQuarter` … 既定の開始月は 4 月
- `YearMonthKey` / `AgeYears`
- `IsJapaneseHoliday` / `JapaneseHolidayName` / `JapaneseHolidaysOfYear` / `WriteJapaneseHolidays`
- `AddWorkDays` / `WorkDaysBetween` / `IsWorkDay`

日本の祝日は 2020-01-01 から 2099-12-31 です。2020 年と 2021 年の東京五輪による移動を含みます。振替休日と国民の休日も返します。春分・秋分は 1980 年から 2099 年向けの近似式なので、官報と違う場合は休業日のセル範囲を `holidays` に渡してください。会社独自の休業日も同じ引数です。

```vba
nextDay = AddWorkDays(DateSerial(2026, 5, 1), 1, , True)
```

2 月 29 日生まれの年齢は、平年では 3 月 1 日を誕生日として数えます。

### modArray / modDict / modTable

- `ToMatrix` / `WriteMatrix` / `MatrixRows` / `MatrixCols` / `ColumnSlice` / `TransposeMatrix` / `WriteCollection`
- `CreateDictionary` / `LoadDictionary` / `DictExists` / `DictItem` / `WriteLookup`
- `HasTable` / `FindTable` / `EnsureTable` / `TableColumnNumber`
- `ClearTableRows` / `AppendTableRow` / `ReadTable` / `FilterTable` / `ClearTableFilter`

辞書のキーは、数値の `1` と文字列の `"1"` を区別します。大文字小文字を無視するかどうかは `compareMode`（0 が区別する、1 が無視する）です。テーブル名はブック内で一意です。

### modFormat（書式・印刷）

- `StyleHeaderRow` / `ApplyGridBorders` / `StripeRows`
- `AutoFitLimited` / `FreezePanesAt` / `FreezeTopRow` / `SetZoom`
- `ApplyNumberFormat` / `FormatAsDate` / `FormatAsInteger` / `FormatAsYen`
- `SetupPrintPage` … A4、ページ数に収める、見出し行を各ページに繰り返す

印刷設定はプリンタと通信するため、初回は時間がかかることがあります。

### modCsv

- `ExportRangeCsv` / `ExportSheetCsv` / `ExportAllSheets` / `ImportCsv`

区切りはカンマです。文字コードは `UTF-8`（BOM 付き）または `Shift_JIS`。日付は `yyyy-mm-dd`、小数点はピリオドで書き出します。引用符、カンマ、セル内改行を扱います。読み込み先より下に残っている古い値は消しません。

### modClean（整形）

- `TrimCells` … 文字列定数だけ。数式は変えない
- `FixNumbersStoredAsText` … `0123` のような先頭ゼロのコードは文字列のまま残す
- `ReplaceText` / `DeduplicateRows`
- `RemoveHyperlinks` / `ClearNotes` / `ClearConditionalFormats`

`DeduplicateRows` は先頭行を見出しとして残し、削除した行数を返します。重複の削除は行全体が対象です。

### modAudit（点検）

- `WriteSheetIndex` / `WriteDefinedNames` / `WriteExternalLinks`
- `WriteHyperlinks` / `WriteFormulaList`

`Run_WriteAuditSheets` は、実行時に開いていたシートの数式を「点検」シートへまとめます。数式は先頭 5000 件までです。

### modForbidden（禁止ワード）

- `FindForbiddenWords` … セル範囲、配列、Collection の一覧と照合し、該当箇所を返す
- `WriteForbiddenWords` … 照合結果をシートへ書く

`Run_FindForbiddenWords` は、禁止ワードの範囲を選ばせてから前面のブックを調べ、結果を「禁止ワード確認」シートへ書きます。一覧に使ったセル自身は該当にしません。セルの値と数式、メモ、コメント、図形、グラフ、ハイパーリンクが対象です。

### modPrivacy（個人情報）

- `FindPersonalData` / `WritePersonalData`

`Run_FindPersonalData` は、前面のブックからメール、電話、郵便番号、カード番号、個人番号に見える文字列を探し、「個人情報確認」シートへ書きます。電話は区切りがあるもの、郵便番号はハイフン付き、カード番号と個人番号は検査数字が合うものだけです。数値セルの連続数字は見ません。メールと番号はマスクします。

### modHidden（非表示データ）

- `FindHiddenContent` / `WriteHiddenContent`

`Run_FindHiddenContent` は、非表示シート、値のある非表示行と列、背景が無いか白いときの白文字、塗りつぶしと同じ色の文字、サイズ 1 または 2 の文字、非表示の名前を「非表示確認」シートへ書きます。濃い背景の白文字は対象にしません。内容列には隠れていたセルの値がそのまま入るので、確認が終わったらこのシートを削除してから提出します。

### modProfile（付帯情報）

- `FindWorkbookProfile` / `WriteWorkbookProfile`

`Run_WriteWorkbookProfile` は、作成者などのプロパティ、ヘッダーとフッター、外部リンク、接続、クエリを「付帯情報」シートへ書きます。接続文字列のパスワード、トークン、API キー、Secret は伏せます。ヘッダーの読み取りには少し時間がかかることがあります。

### modSelfTest

全モジュールを入れた状態で `SelfTest True` を実行すると、一時ブック上で上記の動作を確認します。失敗するとエラーで止まります。確認が終わると一時ブックは閉じます。
