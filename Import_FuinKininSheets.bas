Option Explicit

'================================================================================
' 機能概要（Import_FuinKininSheets_SheetCopy_AndExport）
'--------------------------------------------------------------------------------
' このマクロは「取込」「加工」「更新」「出力」を一気通貫で実行します。
'
' 1) 取込元Excelから対象シートをマクロブックへ取り込み
'    - 固定対象（必須）
'      ・【社保用加工データ】
'         * 完全一致で常に対象（非表示シートでも対象）
'         * シートコピー方式で取り込み
'      ・●支給額
'         * 完全一致で常に対象（非表示でも対象）
'         * ただし「シートコピーしない」
'         * マクロブックに既存の「●支給額」へ値貼りのみ実施
'         * 既存シートを削除・再作成しない（安全要件）
'    - 可変対象
'      ・シート名に「赴任 / 帰任 / 休職」を含むシート
'      ・ただし、以下は除外
'         * 非表示シート
'         * シート名に「マスタ」を含む
'         * シート名に「○」を含む
'      ・対象はシートコピーで取り込み
'
' 2) コピー後の【社保用加工データ】へ赴任情報を追記
'    - 赴任系シートから必要値を抽出し、社保シート末尾へ追加
'    - 追加行に色付け、B列キーで並び替え
'
' 3) 【社保用加工データ】F列の数式を一括再設定
'    - F2に式をセットし最終行までオートフィル
'
' 4) 出力前にリンク更新・RefreshAll・再計算を実施
'
' 5) 【社保用加工データ】を新規ブックに値貼りして保存
'    - B列は表示文字列を再投入して文字列固定
'      （先頭0落ち対策）
'
' 重要仕様
'--------------------------------------------------------------------------------
' - 「●支給額」はマクロブックに事前存在が前提。
' - エラー時であっても「●支給額」は削除対象に含めない。
' - エラー時に削除するのは「今回コピーで新規作成したシート」のみ。
'================================================================================

Public Sub Import_FuinKininSheets_SheetCopy_AndExport()

    '==========================
    ' メッセージ定数（利用者向け）
    '==========================
    Const MSG_CANCEL As String = "ファイル選択がキャンセルされました。"
    Const MSG_NO_TARGET As String = "対象シートが見つかりませんでした。"
    Const MSG_UNPROTECT_FAIL As String = "保護解除に失敗しました。パスワードを確認してください。"
    Const MSG_SUCCESS_FMT As String = "取込が完了しました。（対象：%dシート）"
    Const MSG_NO_SETTING As String = "【設定】シートが見つかりませんでした。"
    Const MSG_EXPORT_EXISTS As String = "出力先ファイルが既に存在します。ファイル名または出力先を変更してください。"
    Const MSG_MISSING_FIXED As String = "取込元ブックに必須シートが存在しません。"
    Const MSG_MISSING_SHIKYU_DST As String = "マクロブックに「●支給額」シートが存在しません。あらかじめ作成してから実行してください。"

    Const PROMPT_BOOK_PW As String = "ブック保護解除パスワードを入力してください"

    '==========================
    ' シート/設定セル 定数
    '==========================
    Const SHEET_SHOHO As String = "【社保用加工データ】"
    Const SHEET_SHIKYU As String = "●支給額"

    Const SETTING_SHEET As String = "【設定】"
    Const SETTING_IN_PATH As String = "E2"
    Const SETTING_OUT_PATH As String = "E4"

    '==========================
    ' ガイド表示文言
    '==========================
    Const GUIDE_PICK_SOURCE As String = "次に、取込元Excelファイルを選択してください。" & vbCrLf & _
                                        "対象：赴任/帰任/休職関連シート、【社保用加工データ】、●支給額 を含むブック"
    Const GUIDE_PICK_OUTFOLDER As String = "次に、出力先フォルダ（ベース）を選択してください。" & vbCrLf & _
                                           "この配下に「YYYYMM」フォルダを作成し、ファイルを書き出します。"
    Const GUIDE_INPUT_YYYYMM As String = "次に、出力対象の年月（YYYYMM）を入力してください。" & vbCrLf & _
                                         "例：202602"

    Const KEY_FUIN As String = "赴任"

    '==========================
    ' ワーク変数
    '==========================
    Dim wbDst As Workbook: Set wbDst = ThisWorkbook          ' 取込先（マクロ格納ブック）
    Dim wsSetting As Worksheet                                ' 設定シート
    Dim wbSrc As Workbook                                     ' 取込元ブック
    Dim srcPath As String                                     ' 取込元ファイルパス

    Dim outBase As String, yyyymm As String                  ' 出力ベースフォルダ / 年月
    Dim outFolder As String, outFilePath As String           ' 実出力フォルダ / 実出力ファイル
    Dim exportSheetName As String                             ' 出力シート名

    ' 実行前のExcelアプリ状態を退避して必ず復元する
    Dim prevSU As Boolean, prevEE As Boolean, prevDA As Boolean
    Dim prevCalc As XlCalculation
    Dim prevStatus As Variant
    Dim wsActiveBefore As Worksheet

    ' ブック構造保護の復帰用
    Dim wasBookProtected As Boolean, bookPw As String

    ' エラー時ロールバック対象（今回新規にコピーしたシートだけ）
    Dim addedSheets As Collection
    Set addedSheets = New Collection

    Dim stepName As String ' どの工程で失敗したかをエラーメッセージに表示
    stepName = "初期化"

    On Error GoTo EH

    '-----------------------------------------
    ' 実行環境退避＆高速化
    '  - ScreenUpdating/EnableEvents/Calculation を抑制
    '  - DisplayAlerts をオフにして削除確認等を抑制
    '-----------------------------------------
    stepName = "実行環境退避＆高速化"
    Set wsActiveBefore = ActiveSheet

    prevSU = Application.ScreenUpdating
    prevEE = Application.EnableEvents
    prevCalc = Application.Calculation
    prevDA = Application.DisplayAlerts
    prevStatus = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = False

    '-----------------------------------------
    ' 【設定】シート存在チェック
    '  - なければ以降の設定保存ができないため終了
    '-----------------------------------------
    stepName = "【設定】シート存在チェック"
    Set wsSetting = GetWorksheetOrNothing(wbDst, SETTING_SHEET)
    If wsSetting Is Nothing Then
        MsgBox MSG_NO_SETTING, vbExclamation
        GoTo SafeExit
    End If

    '-----------------------------------------
    ' 取込元ファイル選択
    '  - ユーザー選択結果を設定シートE2に記録
    '-----------------------------------------
    stepName = "取込元ファイル選択"
    MsgBox GUIDE_PICK_SOURCE, vbInformation, "取込元ファイル選択"
    srcPath = PickOneExcelFileWithTitle("取込元Excelファイルを選択してください")
    If Len(srcPath) = 0 Then
        MsgBox MSG_CANCEL, vbInformation
        GoTo SafeExit
    End If
    wsSetting.Range(SETTING_IN_PATH).Value = srcPath

    '-----------------------------------------
    ' 出力先フォルダ選択
    '  - ユーザー選択結果を設定シートE4に記録
    '-----------------------------------------
    stepName = "出力先フォルダ選択"
    MsgBox GUIDE_PICK_OUTFOLDER, vbInformation, "出力先フォルダ選択"
    outBase = PickOneFolderWithTitle("出力先フォルダを選択してください")
    If Len(outBase) = 0 Then
        MsgBox MSG_CANCEL, vbInformation
        GoTo SafeExit
    End If
    wsSetting.Range(SETTING_OUT_PATH).Value = outBase

    '-----------------------------------------
    ' YYYYMM入力
    '  - 入力値検証は PromptYYYYMMOrBlank で実施
    '-----------------------------------------
    stepName = "YYYYMM入力"
    MsgBox GUIDE_INPUT_YYYYMM, vbInformation, "YYYYMM入力"
    yyyymm = PromptYYYYMMOrBlank("YYYYMM（例：202602）を入力してください")
    If Len(yyyymm) = 0 Then
        MsgBox MSG_CANCEL, vbInformation
        GoTo SafeExit
    End If

    '-----------------------------------------
    ' 取込先ブック保護解除（構造保護）
    '  - 必要時のみ解除
    '  - 後段で必ず元に戻す
    '-----------------------------------------
    stepName = "取込先ブック構造保護解除"
    If Not UnprotectWorkbookStructureIfNeeded(wbDst, PROMPT_BOOK_PW, wasBookProtected, bookPw) Then
        MsgBox MSG_UNPROTECT_FAIL, vbExclamation
        GoTo SafeExit
    End If

    '-----------------------------------------
    ' ●支給額（マクロブック側）存在チェック
    '  - 仕様上「事前に存在」が前提
    '-----------------------------------------
    stepName = "マクロブック：●支給額存在チェック"
    Dim wsShikyuDst As Worksheet
    Set wsShikyuDst = GetWorksheetOrNothing(wbDst, SHEET_SHIKYU)
    If wsShikyuDst Is Nothing Then
        MsgBox MSG_MISSING_SHIKYU_DST, vbExclamation
        GoTo SafeExit
    End If

    '-----------------------------------------
    ' 取込元を ReadOnly で開く
    '-----------------------------------------
    stepName = "取込元ブックを開く"
    Set wbSrc = Workbooks.Open(Filename:=srcPath, ReadOnly:=True)

    '-----------------------------------------
    ' 取込元の必須固定シート存在チェック
    '-----------------------------------------
    stepName = "取込元：必須固定シート存在チェック"
    If GetWorksheetOrNothing(wbSrc, SHEET_SHOHO) Is Nothing Or GetWorksheetOrNothing(wbSrc, SHEET_SHIKYU) Is Nothing Then
        MsgBox MSG_MISSING_FIXED & vbCrLf & _
               "必須：" & SHEET_SHOHO & " / " & SHEET_SHIKYU, vbExclamation
        GoTo SafeExit
    End If

    '-----------------------------------------
    ' 対象シート数カウント
    '  - ●支給額は「値貼り」対象だが、進捗管理のためカウントには含める
    '-----------------------------------------
    stepName = "取込元：対象シート数カウント"
    Dim nTarget As Long
    nTarget = CountTargetSheets(wbSrc, SHEET_SHOHO, SHEET_SHIKYU)
    If nTarget = 0 Then
        MsgBox MSG_NO_TARGET, vbInformation
        GoTo SafeExit
    End If

    '-----------------------------------------
    ' 対象シート取り込み
    '  - 【社保用加工データ】/赴任/帰任/休職：シートコピー
    '  - ●支給額：コピーせず既存シートへ値貼り
    '-----------------------------------------
    stepName = "対象シート取り込み"
    Dim wsSrc As Worksheet
    Dim iTarget As Long
    iTarget = 0

    For Each wsSrc In wbSrc.Worksheets
        If IsTargetSheet(wsSrc, SHEET_SHOHO, SHEET_SHIKYU) Then

            iTarget = iTarget + 1
            Application.StatusBar = "取込中：" & iTarget & "/" & nTarget & " " & wsSrc.Name & "…"

            If StrComp(wsSrc.Name, SHEET_SHIKYU, vbBinaryCompare) = 0 Then
                '-----------------------------------------
                ' ●支給額の取込
                '  - シート削除しない
                '  - シートコピーしない
                '  - 既存シートへ値貼りのみ
                '-----------------------------------------
                stepName = "●支給額：取込元→マクロブックへ値貼り"
                Import_Shikyu_ValuesOnly wsSrc, wsShikyuDst

            Else
                '-----------------------------------------
                ' それ以外の対象はシートコピー
                '  - 既存同名があれば削除後にコピー
                '  - エラー時ロールバック対象として記録
                '-----------------------------------------
                DeleteSheetIfExists wbDst, wsSrc.Name
                wsSrc.Copy After:=wbDst.Worksheets(wbDst.Worksheets.Count)
                addedSheets.Add wsSrc.Name
            End If

        End If
    Next wsSrc

    If iTarget = 0 Then
        MsgBox MSG_NO_TARGET, vbInformation
        GoTo SafeExit
    End If

    '-----------------------------------------
    ' コピー後の社保用加工データに赴任情報を追記
    '-----------------------------------------
    stepName = "社保へ赴任情報追記"
    Application.StatusBar = "社保用加工データへ赴任情報を追記中…"
    AppendFuinInfoToShohoSheet wbDst, SHEET_SHOHO, addedSheets, KEY_FUIN

    '-----------------------------------------
    ' F列数式を全置換（F2～F最終行）
    '-----------------------------------------
    stepName = "社保F列数式置換"
    Application.StatusBar = "【社保用加工データ】のF列数式を更新中…"
    RewriteShohoColumnF_Formulas wbDst, SHEET_SHOHO

    '-----------------------------------------
    ' 出力前にリンク更新＋全更新＋再計算
    '-----------------------------------------
    stepName = "出力前：リンク/更新/再計算"
    Application.StatusBar = "出力前：リンク/接続の更新・再計算中…"
    UpdateLinksAndRefreshAll wbDst

    '-----------------------------------------
    ' 出力処理
    '  - ベース配下に YYYYMM フォルダ作成
    '  - 社保用加工データを値貼り1シートブックとして保存
    '-----------------------------------------
    stepName = "出力：フォルダ作成・ファイル保存"
    Application.StatusBar = "出力準備中…"

    outFolder = JoinPath(outBase, yyyymm)
    EnsureFolderExists outFolder

    exportSheetName = "社保用加工データ_" & yyyymm
    outFilePath = JoinPath(outFolder, exportSheetName & ".xlsx")

    If FileExists(outFilePath) Then
        MsgBox MSG_EXPORT_EXISTS, vbExclamation
        GoTo SafeExit
    End If

    ExportOneSheetAsNewWorkbook_ValueOnly wbDst, SHEET_SHOHO, exportSheetName, outFilePath

    MsgBox Replace(MSG_SUCCESS_FMT, "%d", CStr(iTarget)), vbInformation

SafeExit:
    '-----------------------------------------
    ' 後処理
    '  - 念のためリンク/接続更新
    '  - 開いた取込元ブックをクローズ
    '  - 必要時のみ構造保護を復元
    '  - アプリ状態を実行前に復元
    '-----------------------------------------
    On Error Resume Next
    Application.StatusBar = "終了前：リンク/接続の更新中…"
    UpdateLinksAndRefreshAll wbDst
    On Error GoTo 0

    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0

    On Error Resume Next
    If wasBookProtected Then wbDst.Protect Structure:=True, Password:=bookPw
    On Error GoTo 0

    Application.StatusBar = prevStatus
    Application.DisplayAlerts = prevDA
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEE
    Application.ScreenUpdating = prevSU

    On Error Resume Next
    If Not wsActiveBefore Is Nothing Then wsActiveBefore.Activate
    On Error GoTo 0

    Exit Sub

EH:
    '-----------------------------------------
    ' 途中エラー時のロールバック
    '  - 今回コピー作成したシートのみ削除
    '  - ●支給額は「既存値貼り」運用のため削除対象外
    '-----------------------------------------
    On Error Resume Next

    Dim idx As Long
    For idx = addedSheets.Count To 1 Step -1
        DeleteSheetIfExists wbDst, CStr(addedSheets(idx))
    Next idx

    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    If wasBookProtected Then wbDst.Protect Structure:=True, Password:=bookPw

    Application.StatusBar = prevStatus
    Application.DisplayAlerts = prevDA
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEE
    Application.ScreenUpdating = prevSU

    MsgBox "取込処理中にエラーが発生しました。処理を中止しました。" & vbCrLf & _
           "----" & vbCrLf & _
           "Step: " & stepName & vbCrLf & _
           "Err: " & Err.Number & vbCrLf & _
           "Desc: " & Err.Description, vbExclamation

    On Error GoTo 0
End Sub

'==========================
' 対象シート判定
'--------------------------------------------------------------------------------
' 判定ルール（優先順）：
'  1) 固定2枚（社保 / ●支給額）は「完全一致なら常に対象」
'     - 非表示でも対象
'  2) それ以外は以下を満たすと対象
'     - 可視シート
'     - 「マスタ」を含まない
'     - 「○」を含まない
'     - 「赴任 / 帰任 / 休職」のいずれかを含む
'==========================
Private Function IsTargetSheet(ByVal ws As Worksheet, ByVal shohoName As String, ByVal shikyuName As String) As Boolean

    If StrComp(ws.Name, shohoName, vbBinaryCompare) = 0 Then
        IsTargetSheet = True
        Exit Function
    End If

    If StrComp(ws.Name, shikyuName, vbBinaryCompare) = 0 Then
        IsTargetSheet = True
        Exit Function
    End If

    If ws.Visible <> xlSheetVisible Then Exit Function
    If InStr(1, ws.Name, "マスタ", vbTextCompare) > 0 Then Exit Function
    If InStr(1, ws.Name, "○", vbBinaryCompare) > 0 Then Exit Function

    If InStr(1, ws.Name, "赴任", vbTextCompare) > 0 _
       Or InStr(1, ws.Name, "帰任", vbTextCompare) > 0 _
       Or InStr(1, ws.Name, "休職", vbTextCompare) > 0 Then
        IsTargetSheet = True
    End If
End Function

' 対象シート数カウント（進捗表示用）
Private Function CountTargetSheets(ByVal wb As Workbook, ByVal shohoName As String, ByVal shikyuName As String) As Long
    Dim ws As Worksheet, c As Long
    c = 0
    For Each ws In wb.Worksheets
        If IsTargetSheet(ws, shohoName, shikyuName) Then c = c + 1
    Next ws
    CountTargetSheets = c
End Function

'==========================
' ●支給額：取込元→マクロブック既存シートへ値貼り
'--------------------------------------------------------------------------------
' 目的：
'  - ●支給額シート自体を作り直さず、既存シートへデータのみ反映する。
' 実装方針：
'  - フィルタ表示状態に依存させないため UsedRange 全体を対象
'  - 貼付先は一旦 Cells.Clear（内容・書式クリア）
'  - A1基点で xlPasteValues
' 保護シート対応：
'  - 貼付先が保護されている場合は一時解除を試行
'  - 解除不可なら明示エラーを発生
'==========================
Private Sub Import_Shikyu_ValuesOnly(ByVal wsSrc As Worksheet, ByVal wsDst As Worksheet)

    Dim wasProtected As Boolean
    wasProtected = wsDst.ProtectContents

    If wasProtected Then
        On Error Resume Next
        wsDst.Unprotect Password:=vbNullString
        On Error GoTo 0

        If wsDst.ProtectContents Then
            Err.Raise vbObjectError + 4001, , "●支給額（貼付先）が保護されており値貼りできません。保護解除してください。"
        End If
    End If

    wsDst.Cells.Clear

    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If Not ur Is Nothing Then
        ur.Copy
        wsDst.Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If

    If wasProtected Then
        wsDst.Protect
    End If
End Sub

'==========================
' 社保用加工データへ赴任情報追記
'--------------------------------------------------------------------------------
' 入力：
'  - addedSheets : 今回コピーで作成したシート名一覧
' 処理：
'  1) 赴任キーワードを含むシートのみ抽出
'  2) M2/N2 から年月(YYYYMM)、B4テキスト、B5値を社保シートへ追記
'  3) 既存最終行の D:F 数式を新規行へコピー
'  4) 新規行を薄緑で色付け
'  5) A:G 範囲を B列昇順でソート
'==========================
Private Sub AppendFuinInfoToShohoSheet( _
    ByVal wbDst As Workbook, _
    ByVal shohoSheetName As String, _
    ByVal addedSheets As Collection, _
    ByVal keyFuin As String _
)
    Dim wsShoho As Worksheet
    Set wsShoho = GetWorksheetOrNothing(wbDst, shohoSheetName)
    If wsShoho Is Nothing Then
        Err.Raise vbObjectError + 2101, , "社保用加工データシートが見つかりません: " & shohoSheetName
    End If

    Dim baseLastRow As Long
    baseLastRow = GetLastUsedRowAcross(wsShoho, 1, 3)

    Dim addedRowStart As Long
    addedRowStart = IIf(baseLastRow < 1, 1, baseLastRow + 1)

    Dim addedCount As Long
    addedCount = 0

    Dim i As Long
    For i = 1 To addedSheets.Count
        Dim nm As String
        nm = CStr(addedSheets(i))

        If InStr(1, nm, keyFuin, vbTextCompare) > 0 Then
            Dim ws As Worksheet
            Set ws = GetWorksheetOrNothing(wbDst, nm)
            If Not ws Is Nothing Then

                Dim yyyymm As String
                yyyymm = BuildYYYYMMFromCells(ws.Range("M2").Value, ws.Range("N2").Value)

                Dim b4Text As String
                b4Text = CStr(ws.Range("B4").Text)

                Dim r As Long
                r = addedRowStart + addedCount

                wsShoho.Cells(r, 1).Value = yyyymm
                With wsShoho.Cells(r, 2)
                    .NumberFormat = "@"
                    .Value = b4Text
                End With
                wsShoho.Cells(r, 3).Value = ws.Range("B5").Value

                addedCount = addedCount + 1
            End If
        End If
    Next i

    If addedCount = 0 Then Exit Sub

    Dim newLastRow As Long
    newLastRow = addedRowStart + addedCount - 1

    If baseLastRow >= 1 Then
        wsShoho.Range("D" & baseLastRow & ":F" & baseLastRow).Copy
        wsShoho.Range("D" & (baseLastRow + 1) & ":F" & newLastRow).PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
    End If

    wsShoho.Range("A" & addedRowStart & ":G" & newLastRow).Interior.Color = RGB(237, 251, 219)

    Dim lastRowAll As Long
    lastRowAll = GetLastUsedRowAcross(wsShoho, 1, 7)
    If lastRowAll >= 2 Then
        SortShohoByB wsShoho, lastRowAll
    End If
End Sub

' 社保シートをB列昇順でソート
Private Sub SortShohoByB(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B2:B" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange ws.Range("A1:G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

' 年/月セルから YYYYMM を作る（Mが1桁でも 0 埋め）
Private Function BuildYYYYMMFromCells(ByVal vY As Variant, ByVal vM As Variant) As String
    Dim y As String, m As String
    y = Trim$(CStr(vY))

    If IsNumeric(vM) Then
        m = Format$(CLng(vM), "00")
    Else
        m = Trim$(CStr(vM))
        If Len(m) = 1 And IsNumeric(m) Then m = "0" & m
        If Len(m) = 0 Then m = "00"
    End If

    BuildYYYYMMFromCells = y & m
End Function

' 複数列のうち最も下にある使用行を返す
Private Function GetLastUsedRowAcross(ByVal ws As Worksheet, ByVal colFrom As Long, ByVal colTo As Long) As Long
    Dim c As Long, lastR As Long, r As Long
    lastR = 0
    For c = colFrom To colTo
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If ws.Cells(r, c).Value <> vbNullString Then
            If r > lastR Then lastR = r
        End If
    Next c
    GetLastUsedRowAcross = lastR
End Function

'==========================
' 【社保用加工データ】F列の数式を全置換
'--------------------------------------------------------------------------------
' - B列最終行までを対象
' - F2に基準式を設定し、必要行へオートフィル
'==========================
Private Sub RewriteShohoColumnF_Formulas(ByVal wbDst As Workbook, ByVal shohoSheetName As String)

    Dim wsShoho As Worksheet
    Set wsShoho = GetWorksheetOrNothing(wbDst, shohoSheetName)
    If wsShoho Is Nothing Then
        Err.Raise vbObjectError + 3101, , "社保用加工データシートが見つかりません: " & shohoSheetName
    End If

    Dim lastRow As Long
    lastRow = GetLastRowInColumn(wsShoho, 2)
    If lastRow < 2 Then Exit Sub

    Dim f As String
    f = "=IF(" & _
        "XLOOKUP(B2,'(追加)日割・国内支給者'!D:D,'(追加)日割・国内支給者'!J:J,"""")=""国内支給""," & _
        "0," & _
        "XLOOKUP(B2,'(追加)日割・国内支給者'!D:D,'(追加)日割・国内支給者'!I:I," & _
            "VLOOKUP(B2,●支給額!$B:$CJ,83,FALSE)" & _
        ")" & _
    ")"

    With wsShoho
        .Range("F2").Formula = f
        If lastRow >= 3 Then
            .Range("F2").AutoFill Destination:=.Range("F2:F" & lastRow), Type:=xlFillDefault
        End If
    End With
End Sub

'==========================
' リンク更新＋全更新＋再計算
'--------------------------------------------------------------------------------
' 失敗しても止めない（On Error Resume Next）
'  - 外部リンクが無いブックでも安全に通すため
'==========================
Private Sub UpdateLinksAndRefreshAll(ByVal wb As Workbook)
    On Error Resume Next
    Dim arr As Variant
    arr = wb.LinkSources(Type:=xlExcelLinks)
    If Not IsEmpty(arr) Then
        wb.UpdateLink Name:=arr, Type:=xlExcelLinks
    End If
    wb.RefreshAll
    Application.Calculate
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
End Sub

'==========================
' 出力：1シートを新規ブックとして保存（値固定）
'--------------------------------------------------------------------------------
' B列先頭ゼロ落ち対策：
'  1) 先にB列の表示文字列を配列へ退避
'  2) UsedRangeを値化（式を除去）
'  3) B列へ文字列として書き戻し（NumberFormat="@"）
'==========================
Private Sub ExportOneSheetAsNewWorkbook_ValueOnly( _
    ByVal wbSrc As Workbook, _
    ByVal srcSheetName As String, _
    ByVal newSheetName As String, _
    ByVal outFullPath As String _
)
    Dim ws As Worksheet
    Set ws = GetWorksheetOrNothing(wbSrc, srcSheetName)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2001, , "Export sheet not found: " & srcSheetName
    End If

    ws.Copy

    Dim wbNew As Workbook
    Set wbNew = ActiveWorkbook

    Dim wsNew As Worksheet
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = newSheetName

    Dim bLast As Long
    bLast = GetLastRowInColumn(wsNew, 2)
    If bLast < 2 Then bLast = 2

    Dim bText() As String
    bText = CaptureDisplayedText(wsNew, 2, 2, bLast)

    Dim ur As Range
    Set ur = wsNew.UsedRange
    If Not ur Is Nothing Then
        ur.Value = ur.Value
    End If

    Dim r As Long, i As Long
    i = 0
    For r = 2 To bLast
        i = i + 1
        With wsNew.Cells(r, 2)
            .NumberFormat = "@"
            .Value = bText(i)
        End With
    Next r

    wbNew.SaveAs Filename:=outFullPath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
End Sub

' 指定範囲のセル表示文字列（.Text）を配列で取得
Private Function CaptureDisplayedText( _
    ByVal ws As Worksheet, _
    ByVal col As Long, _
    ByVal rowFrom As Long, _
    ByVal rowTo As Long _
) As String()
    Dim n As Long
    n = rowTo - rowFrom + 1

    Dim arr() As String
    If n < 1 Then
        ReDim arr(1 To 1)
        arr(1) = vbNullString
        CaptureDisplayedText = arr
        Exit Function
    End If

    ReDim arr(1 To n)

    Dim r As Long, i As Long
    i = 0
    For r = rowFrom To rowTo
        i = i + 1
        arr(i) = CStr(ws.Cells(r, col).Text)
    Next r

    CaptureDisplayedText = arr
End Function

' 単一列の最終行を取得（最低1）
Private Function GetLastRowInColumn(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r < 1 Then r = 1
    GetLastRowInColumn = r
End Function

'==========================
' YYYYMM入力（不正なら空）
'--------------------------------------------------------------------------------
' 妥当条件：
'  - 6桁数字
'  - 末尾2桁（月）が 01～12
'==========================
Private Function PromptYYYYMMOrBlank(ByVal prompt As String) As String
    Dim s As String
    s = InputBox(prompt, "YYYYMM入力")
    s = Trim$(s)

    If Len(s) = 0 Then
        PromptYYYYMMOrBlank = vbNullString
        Exit Function
    End If
    If Len(s) <> 6 Then
        PromptYYYYMMOrBlank = vbNullString
        Exit Function
    End If
    If Not IsNumeric(s) Then
        PromptYYYYMMOrBlank = vbNullString
        Exit Function
    End If

    Dim mm As Long
    mm = CLng(Right$(s, 2))
    If mm < 1 Or mm > 12 Then
        PromptYYYYMMOrBlank = vbNullString
        Exit Function
    End If

    PromptYYYYMMOrBlank = s
End Function

' ファイル選択ダイアログ（単一選択）
Private Function PickOneExcelFileWithTitle(ByVal titleText As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = titleText
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls", 1
        If .Show <> -1 Then
            PickOneExcelFileWithTitle = vbNullString
            Exit Function
        End If
        PickOneExcelFileWithTitle = .SelectedItems(1)
    End With
End Function

' フォルダ選択ダイアログ（単一選択）
Private Function PickOneFolderWithTitle(ByVal titleText As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = titleText
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickOneFolderWithTitle = vbNullString
            Exit Function
        End If
        PickOneFolderWithTitle = .SelectedItems(1)
    End With
End Function

' 指定名Worksheetを返す（未存在時は Nothing）
Private Function GetWorksheetOrNothing(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

' 同名シートがあれば削除（呼び出し側で DisplayAlerts=False 前提）
Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = GetWorksheetOrNothing(wb, sheetName)
    If Not ws Is Nothing Then ws.Delete
End Sub

' ブック構造保護の解除（必要時のみ）
'  - 空パスワード解除を先に試行
'  - 失敗時はInputBoxで入力を促して再試行
'  - 成否と利用パスワードを返して、最後に再保護できるようにする
Private Function UnprotectWorkbookStructureIfNeeded( _
    ByVal wb As Workbook, _
    ByVal prompt As String, _
    ByRef wasProtected As Boolean, _
    ByRef usedPassword As String _
) As Boolean

    wasProtected = False
    usedPassword = vbNullString
    UnprotectWorkbookStructureIfNeeded = True

    If wb.ProtectStructure Then
        wasProtected = True

        On Error Resume Next
        wb.Unprotect Password:=vbNullString
        On Error GoTo 0

        If wb.ProtectStructure Then
            usedPassword = InputBox(prompt, "パスワード入力")
            If Len(usedPassword) = 0 Then
                UnprotectWorkbookStructureIfNeeded = False
                Exit Function
            End If

            On Error Resume Next
            wb.Unprotect Password:=usedPassword
            On Error GoTo 0

            If wb.ProtectStructure Then
                UnprotectWorkbookStructureIfNeeded = False
                Exit Function
            End If
        End If
    End If
End Function

' パス結合（末尾の "\" 重複を吸収）
Private Function JoinPath(ByVal basePath As String, ByVal child As String) As String
    If Len(basePath) = 0 Then
        JoinPath = child
        Exit Function
    End If

    If Right$(basePath, 1) = "\" Then
        JoinPath = basePath & child
    Else
        JoinPath = basePath & "\" & child
    End If
End Function

' フォルダ作成（未存在時のみ作成）
Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) <> vbNullString Then Exit Sub
    MkDir folderPath
End Sub

' ファイル存在確認
Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
End Function
