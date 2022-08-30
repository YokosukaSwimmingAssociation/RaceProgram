Attribute VB_Name = "WinnerReportModule"
Option Explicit    ''←変数の宣言を強制する

Public Sub 優勝者一覧作成()

    ' イベント発生を抑制
    Call EventChange(False)

    Dim sGameName As String
    sGameName = GetRange("大会名").Value

    Dim oWinnerList As Object
    Set oWinnerList = CreateObject("Scripting.Dictionary")

    ' 優勝者の読込み
    Call ReadWinner(sGameName, oWinnerList)
    
    ' 優勝者の書込み
    Call WriteWinner(sGameName, oWinnerList)

    ' イベント発生を発生
    Call EventChange(True)

    ' 保存
    ActiveWorkbook.Save
End Sub

'
' 優勝者読込み
'
' sGameName     IN  大会名
' oWinnerList   OUT 優勝者リスト
'
' oWinnerList
' │
' └─種目番号&区分
' 　　│
' 　　└─インデックス：
' 　　　　│
' 　　　　├─氏名
' 　　　　│
' 　　　　├─所属
' 　　　　│
' 　　　　├─区分
' 　　　　│
' 　　　　├─記録
' 　　　　│
' 　　　　└─大会新
'
'
Private Sub ReadWinner(sGameName As String, oWinnerList As Object)
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' プログラム番号毎
    Dim vProNo As Range
    For Each vProNo In GetAreaKeyData(sMasterName)

        ' 決勝（タイム決勝）の場合
        If Not IsSenshukenQualifyRace(sGameName, vProNo) Then
            ' プログラム番号内から１位を探す
            Dim vCell As Range
            For Each vCell In GetRange("プログラム番号" & CStr(vProNo))
                ' １位の場合
                If GetOffset(vCell, Range("Header順位").Column).Value = 1 Then
                    ' 優勝者情報を登録
                    Call SetWinnerInfo(sGameName, vProNo, oWinnerList, vCell)
                End If
            Next vCell
        End If
    Next vProNo
End Sub

'
' 選手権の予選種目か判定
'
' True: 予選／ False: 決勝または選手権以外
'
' sGameName     IN          大会名
' vProNo        IN          ProNo
'
Private Function IsSenshukenQualifyRace(sGameName As String, vProNo As Variant) As Boolean
    ' 選手権の予選は飛ばす
    If sGameName = 選手権大会 Then
         If VLookupArea(vProNo, "選手権種目区分", "予選／決勝") = "予選" Then
            IsSenshukenQualifyRace = True
         Else
            IsSenshukenQualifyRace = False
         End If
    Else
        IsSenshukenQualifyRace = False
    End If
End Function

'
' 優勝者の選手情報登録
'
' sGameName     IN          大会名
' vProNo        IN          ProNo
' oWinnerList   IN/OUT      優勝者配列
' vCell         IN          優勝者のセル
'
Private Sub SetWinnerInfo(sGameName As String, vProNo As Variant, oWinnerList As Object, vCell As Range)
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' １位リスト
    Dim oWinners As Object
    Dim oWinner As Object
    Set oWinner = CreateObject("Scripting.Dictionary")
    
    ' 大会記録が空欄の場合がるのでVariantで宣言
    Dim nRecord As Variant
    nRecord = GetOffset(vCell, Range("Header大会記録").Column).Value
    
    Dim sKey As String
    sKey = GetRecordKey(sGameName, CInt(vProNo), _
        GetOffset(vCell, Range("Header区分").Column).Value)
    
    oWinner.Add "氏名", GetOffset(vCell, Range("Header氏名").Column).Value
    oWinner.Add "所属", GetOffset(vCell, Range("Header所属").Column).Value
    oWinner.Add "記録", GetOffset(vCell, Range("Header時間").Column).Value
    oWinner.Add "区分", GetOffset(vCell, Range("Header区分").Column).Value
    
    If Not IsNumeric(nRecord) Then
        oWinner.Add "大会新", "参考記録"
    ElseIf oWinner.Item("記録") <= nRecord Then
        oWinner.Add "大会新", "大会新"
    End If
    
    ' プロNo＋区分の１位が未登録の場合
    If Not (oWinnerList.Exists(sKey)) Then
        Set oWinners = CreateObject("Scripting.Dictionary")
        oWinners.Add oWinners.Count + 1, oWinner
        oWinnerList.Add sKey, oWinners
    Else
        Set oWinners = oWinnerList.Item(sKey)
        oWinners.Add oWinners.Count + 1, oWinner
    End If
End Sub

'
' 大会毎の大会記録の区分
'
' sGameName     IN  大会名
' nProNo        IN  種目番号
' sClass        IN  区分
'
Public Function GetRecordKey(sGameName As String, nProNo As Integer, sClass As String) As String

    If sGameName = 選手権大会 Then
        GetRecordKey = CStr(nProNo)
    ElseIf sGameName = 市民大会 Then
        GetRecordKey = Format(nProNo, "00") & "_" & Replace(STrimAll(sClass), "一般", "20代")
    Else
        Dim sMasterName As String
        sMasterName = GetMaster(sGameName)
        ' 区分を取得
        If Trim(VLookupArea(nProNo, sMasterName, "種目区分")) = "" Then
            GetRecordKey = CStr(nProNo) & "_" & sClass
        Else
            GetRecordKey = CStr(nProNo) & "_"
        End If
    End If

End Function

'
' 優勝者書込み
'
' sGameName     IN  大会名
' oWinnerList   IN  優勝者リスト
'
Private Sub WriteWinner(sGameName As String, oWinnerList As Object)

    ' 優勝者シートを選択し保護を解除
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)

    ' 優勝者範囲名
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
        
    If sGameName = 選手権大会 Then
        ' 初期化
        Call DeleteWinnerSheetForSenshuken(sWinnerAreaName)
        
        ' 出力する
        Call WriteWinnerListForSenshuken(sWinnerAreaName, oWinnerList)
    
        ' 優勝者シート設定
        Call 選手権大会記録設定(sSheetName, sWinnerAreaName)
    Else
        ' 初期化
        Call DeleteWinnerSheet(sWinnerAreaName)
        
        ' 出力する
        Call WriteWinnerList(sWinnerAreaName, sRecordAreaName, oWinnerList)
        
        ' 優勝者シート設定
        Call DefineWinnerSheet(sSheetName, sWinnerAreaName)
    
        ' 書式設定
        Call SetWinnerRecordStyle(sGameName)
    End If

    ' シートの保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 優勝者書込み
'
' sWinnerAreaName   IN  優勝者範囲名
' sRecordAreaName   IN  大会記録範囲名
' oWinnerList       IN  優勝者リスト
'
Private Sub WriteWinnerList(sWinnerAreaName As String, sRecordAreaName As String, oWinnerList As Object)
    
    ' 同一の区分に対して１回のみ実施するチェック用
    Dim oKeyList As Object
    Set oKeyList = CreateObject("Scripting.Dictionary")
    
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
        
    ' 大会記録毎
    Dim vKey As Range
    For Each vKey In GetAreaKeyData(sRecordAreaName)
        If Not oKeyList.Exists(CStr(vKey.Value)) Then
            If oWinnerList.Exists(CStr(vKey.Value)) Then
                Set oWinners = oWinnerList.Item(CStr(vKey.Value))
                Dim vIdx As Variant
                For Each vIdx In oWinners
                    Set oWinner = oWinners.Item(vIdx)
                    Call WriteWinnerLine(sWinnerAreaName, sRecordAreaName, nRow, vKey, oWinner)
                    nRow = nRow + 1
                Next vIdx
            End If
            oKeyList.Add CStr(vKey.Value), 1
        End If
    Next vKey
End Sub

'
' 優勝者書込み（選手権用）
'
' sAreaName         IN  優勝者範囲名
' oWinnerList       IN  優勝者リスト
'
Private Sub WriteWinnerListForSenshuken(sAreaName As String, oWinnerList As Object)
    
    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")
    
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData(sAreaName)
        If vProNo.Value > 0 Then
            ' 優勝者が存在する場合
            If oWinnerList.Exists(CStr(vProNo.Value)) Then
                ' 同一ProNoには１回だけ実施
                If Not oProNo.Exists(CStr(vProNo.Value)) Then
                    Set oWinners = oWinnerList.Item(CStr(vProNo.Value))
                    Dim vIdx As Variant
                    For Each vIdx In oWinners
                        Set oWinner = oWinners.Item(vIdx)
                        If vIdx > 1 Then
                            ' 同タイム記録の場合は行を挿入
                            Set vProNo = InsertWinnerRow(vProNo, sAreaName)
                        End If
                        Call WriteWinnerLineForSenshuken(sAreaName, vProNo, oWinner)
                    Next vIdx
                    oProNo.Add CStr(vProNo.Value), 1
                End If
            End If
        End If
    Next vProNo

End Sub

'
' 優勝者シート初期化
'
Public Sub 優勝者シート初期化()
    Call DeleteWinnerSheet(GetWinnerSheetName(GetRange("大会名").Value))
End Sub

'
' 優勝者シート初期化
'
' sWinnerAreaName   IN  優勝者範囲名
' nRow              IN  優勝者シートを消す行
'
Private Sub DeleteWinnerSheet(sWinnerAreaName As String, Optional nRow As Integer = 1)
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    If Cells(oRange.Row + nRow, oRange.Column) <> "" Then
        oRange.Offset(nRow, 0).Resize(oRange.Rows().Count - nRow).EntireRow.Delete
    End If
End Sub

'
' 優勝者シート初期化（選手権用）
'
' sAreaName         IN  優勝者範囲名
'
Private Sub DeleteWinnerSheetForSenshuken(sAreaName As String)

    ' 複数行ある場合は削除最小とする
    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")
    Dim oDelList As Object
    Set oDelList = CreateObject("Scripting.Dictionary")

    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData(sAreaName)
        If vProNo.Value > 0 Then
            If oProNo.Exists(vProNo.Value) Then
                ' ループ中に削除すると範囲がおかしくなるので
                ' 一旦削除対象リストに追加する
                oDelList.Add oDelList.Count + 1, vProNo
            Else
                oProNo.Add vProNo.Value, 1
                GetOffset(vProNo, GetColIdx(sAreaName, "氏名")) = ""
                GetOffset(vProNo, GetColIdx(sAreaName, "所属")) = ""
                GetOffset(vProNo, GetColIdx(sAreaName, "区分")) = ""
                GetOffset(vProNo, GetColIdx(sAreaName, "記録")) = ""
                If GetColIdx(sAreaName, "大会新") > 0 Then
                    GetOffset(vProNo, GetColIdx(sAreaName, "大会新")) = ""
                End If
                If GetColIdx(sAreaName, "年") > 0 Then
                    GetOffset(vProNo, GetColIdx(sAreaName, "年")) = ""
                End If
            End If
        End If
    Next vProNo

    ' 削除リストを削除する
    Dim oDelCell As Range
    Dim vDelCell As Variant
    For Each vDelCell In oDelList
        Set oDelCell = oDelList.Item(vDelCell)
        oDelCell.EntireRow.Delete
    Next vDelCell

End Sub

'
' 優勝者シート記入
'
' sWinnerAreaName   IN  優勝者範囲名
' sRecordAreaName   IN  大会記録範囲名
' nRow              IN  優勝者の行数
' vKey              IN  大会記録の参照元の基準セル
' oWinner           IN  優勝者情報
'
Private Sub WriteWinnerLine(sWinnerAreaName As String, sRecordAreaName As String, _
nRow As Integer, vKey As Variant, oWinner As Object)
    
    Cells(nRow, GetColIdx(sWinnerAreaName, "プロNo.")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "プロNo."))
    Cells(nRow, GetColIdx(sWinnerAreaName, "性別")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "性別"))
    Cells(nRow, GetColIdx(sWinnerAreaName, "距離")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "距離"))
    Cells(nRow, GetColIdx(sWinnerAreaName, "種目")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "種目"))
    Cells(nRow, GetColIdx(sWinnerAreaName, "区分")) = GetOffset(vKey, GetColIdx(sRecordAreaName, "区分"))
    
    Cells(nRow, GetColIdx(sWinnerAreaName, "氏名")) = oWinner.Item("氏名")
    Cells(nRow, GetColIdx(sWinnerAreaName, "所属")) = oWinner.Item("所属")
    Cells(nRow, GetColIdx(sWinnerAreaName, "記録")) = oWinner.Item("記録")
    Cells(nRow, GetColIdx(sWinnerAreaName, "大会新")) = oWinner.Item("大会新")

End Sub

'
' 優勝者シート記入（選手権用）
'
' sAreaName         IN  優勝者範囲名
' vKey              IN  優勝者の基準セル
' oWinner           IN  優勝者情報
'
Private Sub WriteWinnerLineForSenshuken(sAreaName As String, vKey As Variant, oWinner As Object)
    
    GetOffset(vKey, GetColIdx(sAreaName, "氏名")) = Replace(oWinner.Item("氏名"), "．", vbCrLf)
    GetOffset(vKey, GetColIdx(sAreaName, "所属")) = oWinner.Item("所属")
    GetOffset(vKey, GetColIdx(sAreaName, "区分")) = oWinner.Item("区分")
    GetOffset(vKey, GetColIdx(sAreaName, "記録")) = oWinner.Item("記録")
    
    If GetColIdx(sAreaName, "大会新") > 0 Then
        GetOffset(vKey, GetColIdx(sAreaName, "大会新")) = oWinner.Item("大会新")
    End If
    
    If GetColIdx(sAreaName, "年") > 0 Then
        GetOffset(vKey, GetColIdx(sAreaName, "年")) = oWinner.Item("年")
    End If
    

End Sub

'
' 優勝者シート行挿入（選手権用）
'
' 同タイムの場合に行を挿入する
'
' oCell         IN      基準のセル
' sAreaName         IN  優勝者範囲名
'
Private Function InsertWinnerRow(oCell As Variant, sAreaName As String) As Range
    Dim oNewCell As Range
    oCell.Offset(1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Set oNewCell = oCell.Offset(1)
    oNewCell.Value = oCell.Value
    GetOffset(oNewCell, GetColIdx(sAreaName, "性別")).Value = GetOffset(oCell, GetColIdx(sAreaName, "性別")).Value
    ' 種目列は結合
    Union(GetOffset(oNewCell, GetColIdx(sAreaName, "種目")), GetOffset(oCell, GetColIdx(sAreaName, "種目")).MergeArea).Merge
    GetOffset(oNewCell, GetColIdx(sAreaName, "距離")).Value = GetOffset(oCell, GetColIdx(sAreaName, "距離")).Value
    ' 全体に罫線を設定
    Call SetBorder(GetOffset(oNewCell, GetColIdx(sAreaName, "種目")) _
        .Resize(1, GetRange(sAreaName).Columns.Count - (GetColIdx(sAreaName, "種目") - oCell.Column)))
    ' 新しい行の基準セルを返す
    Set InsertWinnerRow = oNewCell
End Function


'
' 優勝者シートの書式を大会記録からコピーする
'
' sGameName         IN  大会名
'
Private Sub SetWinnerRecordStyle(sGameName As String)

    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    
     ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    GetRange(sRecordAreaName).Offset(1, 1).Resize(1).Copy
        
    ' 優勝者シート
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    
    ' 優勝者範囲名
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    If Cells(oRange.Row + 1, oRange.Column) <> "" Then
        oRange.Offset(1, 0).Resize(oRange.Rows.Count - 1, oRange.Columns.Count).Select
    End If
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Range("A1").Select
End Sub

'
' 大会記録シートの書式を２行目からコピーする
'
' sGameName         IN  大会名
' sSheetName        IN  シート名
' sAreaName         IN  範囲名
'
Private Sub SetRecordWinnerStyle(sGameName As String, sSheetName As String, sAreaName As String)

    Dim nFormatRow As Integer
    nFormatRow = 1
    Dim nStartRow As Integer
    nStartRow = 2

    ' シートをアクティブ化
    Sheets(sSheetName).Activate
    
    ' 範囲名の2行目（データ１行目を選択）
    GetRange(sAreaName).Offset(nFormatRow, 0).Resize(1).Copy
    
    Dim oRange As Range
    Set oRange = TableRange(sAreaName)
    If oRange.Offset(nStartRow).Resize(1, 1).Value <> "" Then
        oRange.Offset(nStartRow, 0).Resize(oRange.Rows.Count - nStartRow).Select
    End If
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Range("A1").Select
End Sub

'
' 大会記録シートを並び替える
'
' sSheetName        IN  シート名
' sAreaName         IN  範囲名
'
Private Sub SortRecordWinner(sSheetName As String, sAreaName As String)

    ' シートをアクティブ化
    Sheets(sSheetName).Activate
    Range("A1").Select
    Selection.AutoFilter
   
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:= _
            Range(RowRangeAddress("B2")), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
            :=xlSortTextAsNumbers
        .SortFields.Add2 Key:= _
            Range(RowRangeAddress("F2")), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
            :=xlSortNormal

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A1").Select
End Sub

'
' 大会記録シートを並び替える（市民大会）
'
' sSheetName        IN  シート名
' sAreaName         IN  範囲名
'
Private Sub SortRecordWinnerShimin(sSheetName As String, sAreaName As String)

    ' シートをアクティブ化
    Sheets(sSheetName).Activate
    Range("A1").Select
    Selection.AutoFilter
   
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:= _
            Range(RowRangeAddress("A2")), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
            :=xlSortNormal

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A1").Select
End Sub

'
' 優勝者シート名
'
' sGameName     IN      大会名
'
Private Function GetWinnerSheetName(sGameName As String) As String
    GetWinnerSheetName = VLookupArea(sGameName, "設定各種", "優勝者シート名")
End Function

'
' 優勝者範囲名
'
' sGameName     IN      大会名
'
Private Function GetWinnerAreaName(sGameName As String) As String
    GetWinnerAreaName = VLookupArea(sGameName, "設定各種", "優勝者範囲名")
End Function

'
' 大会記録シート名
'
' sGameName     IN      大会名
'
Private Function GetRecordSheetName(sGameName As String) As String
    GetRecordSheetName = VLookupArea(sGameName, "設定各種", "大会記録シート名")
End Function


'
' 大会記録範囲名
'
' sGameName     IN      大会名
'
Private Function GetRecordAreaName(sGameName As String) As String
    GetRecordAreaName = VLookupArea(sGameName, "設定各種", "大会記録範囲名")
End Function

'
' 大会記録更新
'
Public Sub 大会記録更新()

    ' イベント発生を抑制
    Call EventChange(False)

    Dim sGameName As String
    sGameName = GetRange("大会名").Value

    Dim oWinnerList As Object
    Set oWinnerList = CreateObject("Scripting.Dictionary")

    ' 大会記録の読込み
    Call ReadRecords(sGameName, oWinnerList)

    ' 新大会記録の読込み
    Call ReadNewRecords(sGameName, oWinnerList)
    
    If sGameName = 選手権大会 Then
        ' 大会記録の書込み
        Call WriteNewRecordsForSenshuken(sGameName, oWinnerList)
    Else
        ' 大会記録の書込み
        Call WriteNewRecords(sGameName, oWinnerList)
    End If

    ' イベント発生を発生
    Call EventChange(True)

    ' 保存
    ActiveWorkbook.Save

    ' 大会記録シート
    Call SheetActivate(GetRecordSheetName(sGameName))

End Sub

'
' 大会記録読込み
'
' sGameName     IN  大会名
' oRecordList   OUT 大会記録者リスト
'
' oRecordList
' │
' └─種目番号&区分
' 　　│
' 　　└─インデックス：
' 　　　　│
' 　　　　├─氏名
' 　　　　│
' 　　　　├─所属
' 　　　　│
' 　　　　├─記録
' 　　　　│
' 　　　　└─大会新
'
'
Private Sub ReadRecords(sGameName As String, oRecordList As Object)
    
    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    
    Call SheetActivate(sSheetName)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    ' 大会年
    Dim nYear As Integer
    nYear = GetRange("大会年").Value
    
    Dim sKey As String
        
    ' 大会記録毎
    Dim vCell As Range
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        sKey = vCell.Value
        
        Set oWinner = CreateObject("Scripting.Dictionary")
        
        ' カラムの値の登録
        Dim vKey As Range
        For Each vKey In GetRange(sRecordAreaName).Rows(1).Columns()
            oWinner.Add STrimAll(vKey.Value), GetOffset(vCell, vKey.Column).Value
        Next vKey
        
        ' 大会年と異なる場合読み込む
        If oWinner.Item("年") <> nYear Then
            ' プロNo＋区分の１位が未登録の場合
            If Not (oRecordList.Exists(sKey)) Then
                Set oWinners = CreateObject("Scripting.Dictionary")
                oWinners.Add oWinners.Count + 1, oWinner
                oRecordList.Add sKey, oWinners
            Else
                Set oWinners = oRecordList.Item(sKey)
                oWinners.Add oWinners.Count + 1, oWinner
            End If
        End If
    Next vCell
End Sub

'
' 大会記録読込み
'
' sGameName     IN  大会名
' oRecordList   OUT 大会記録者リスト
'
' oRecordList
' │
' └─種目番号&区分
' 　　│
' 　　└─インデックス：
' 　　　　│
' 　　　　├─氏名
' 　　　　│
' 　　　　├─所属
' 　　　　│
' 　　　　├─記録
' 　　　　│
' 　　　　└─大会新
'
'
Private Sub ReadNewRecords(sGameName As String, oRecordList As Object)
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' 優勝者シート
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select

    ' 優勝者の範囲
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    
    ' 大会記録の範囲
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    ' 大会年
    Dim nYear As Integer
    nYear = GetRange("大会年").Value
    
    Dim oWinners As Object
    Dim oWinner As Object
    Dim oWinnerOld As Object

    Dim sKeyName As String
    Dim sKey As String

    ' 優勝者毎
    Dim vCell As Range
    For Each vCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' 大会新なら格納する
        If GetOffset(vCell, GetColIdx(sWinnerAreaName, "大会新")).Value = "大会新" _
            Or GetOffset(vCell, GetColIdx(sWinnerAreaName, "大会新")).Value = "参考記録" Then
                
            Set oWinner = CreateObject("Scripting.Dictionary")
            
            ' キーを取得
            sKeyName = GetAreaKeyName(sRecordAreaName)
            sKey = GetRecordKey(sGameName, CInt(vCell.Value), _
                    GetOffset(vCell, GetColIdx(sWinnerAreaName, "区分")))
            
            ' 優勝者の列値をすべて登録
            Dim vKey As Variant
            For Each vKey In GetRange(sWinnerAreaName).Rows(1).Columns()
                oWinner.Add STrimAll(vKey.Value), GetOffset(vCell, vKey.Column).Value
            Next vKey
            
            ' 選手権の場合はKeyが範囲内にあるのでチェックしてから追加
            If Not oWinner.Exists(sKeyName) Then
                oWinner.Add sKeyName, sKey
            End If
            oWinner.Add "年", nYear
                
            ' プロNo＋区分の１位が未登録の場合
            If Not (oRecordList.Exists(sKey)) Then
                Set oWinners = CreateObject("Scripting.Dictionary")
                oWinners.Add oWinners.Count + 1, oWinner
                oRecordList.Add sKey, oWinners
            Else
                Set oWinners = oRecordList.Item(sKey)
                ' 既に存在する場合はタイムを比較し古ければ削除
                ' 過去の大会新が存在しない場合も削除する
                For Each vKey In oWinners.Keys()
                    Set oWinnerOld = oWinners.Item(vKey)
                    If oWinner.Item("記録") < oWinnerOld.Item("記録") Or _
                       oWinnerOld.Item("記録") = "" Then
                        oWinners.Remove vKey
                    End If
                Next vKey
                
                ' 追加する
                oWinners.Add oWinners.Count + 1, oWinner
            End If
        End If
    Next vCell
End Sub

'
' 大会記録書込み
'
' sGameName     IN  大会名
' oWinnerList   IN  優勝者リスト
'
Private Sub WriteNewRecords(sGameName As String, oRecordList As Object)

    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    ' フィルタ解除
    Call SetAutoFilter(sRecordAreaName, False)

    ' 削除
    Call DeleteWinnerSheet(sRecordAreaName, 2)

    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
    
    ' 大会記録毎
    Dim vKey As Variant
    For Each vKey In oRecordList.Keys()
        Set oWinners = oRecordList.Item(vKey)
        Dim vIdx As Variant
        For Each vIdx In oWinners
            Set oWinner = oWinners.Item(vIdx)
            
            ' カラムの値の書込み
            Dim vCell As Range
            For Each vCell In GetRange(sRecordAreaName).Rows(1).Columns()
                Cells(nRow, vCell.Column) = oWinner.Item(STrimAll(vCell.Value))
            Next vCell
            
            nRow = nRow + 1
        Next vIdx
    Next vKey

    ' 書式設定
    Call SetRecordWinnerStyle(sGameName, sSheetName, sRecordAreaName)

    ' 並び替え
    If sGameName = 市民大会 Then
        Call SortRecordWinnerShimin(sSheetName, sRecordAreaName)
    Else
        Call SortRecordWinner(sSheetName, sRecordAreaName)
    End If

    ' 大会記録シートの設定
    Call DefineRecordSheet(sSheetName, sRecordAreaName)

    ' シートの保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 大会記録書込み（選手権用）
'
' sGameName     IN  大会名
' oWinnerList   IN  優勝者リスト
'
Private Sub WriteNewRecordsForSenshuken(sGameName As String, oRecordList As Object)

    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    
    ' 表示／アクティブ／解除
    Dim vVisible As Variant
    vVisible = GetSheetVisible(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(sSheetName)
    Call SheetProtect(False, oWorkSheet)
    
    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    ' 初期化
    Call DeleteWinnerSheetForSenshuken(sRecordAreaName)

    ' 書込み
    Call WriteWinnerListForSenshuken(sRecordAreaName, oRecordList)

    ' 優勝者シート設定
    Call 選手権大会記録設定(sSheetName, sRecordAreaName)

    ' シートの保護
    Call SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = vVisible
End Sub

'
' 大会記録シート記入
'
' sGameName     IN  大会名
' vCell         IN  参照元の基準セル
' oWinner       IN  優勝者情報
'
Private Sub WriteRecordLine(sAreaName As String, vCell As Variant, oWinner As Object)

    vCell.Offset(0, GetColIdx(sAreaName, "氏名") - 1).Value = oWinner.Item("氏名")
    vCell.Offset(0, GetColIdx(sAreaName, "所属") - 1).Value = oWinner.Item("所属")
    vCell.Offset(0, GetColIdx(sAreaName, "記録") - 1).Value = oWinner.Item("記録")
    vCell.Offset(0, GetColIdx(sAreaName, "年") - 1).Value = oWinner.Item("年")

End Sub
