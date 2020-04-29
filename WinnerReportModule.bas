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
    sKey = GetWinnerKey(sGameName, sMasterName, CInt(vProNo), _
        GetOffset(vCell, Range("Header区分").Column).Value)
    
    oWinner.Add "氏名", GetOffset(vCell, Range("Header氏名").Column).Value
    oWinner.Add "所属", GetOffset(vCell, Range("Header所属").Column).Value
    oWinner.Add "記録", GetOffset(vCell, Range("Header時間").Column).Value
    
    If Not IsNumeric(nRecord) Or oWinner.Item("記録") <= nRecord Then
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
' sMasterName   IN  マスター名
' nProNo        IN  種目番号
' sType         IN  区分
'
Private Function GetWinnerKey(sGameName As String, sMasterName As String, nProNo As Integer, sType As String)

    If sGameName = 選手権大会 Then
        GetWinnerKey = CStr(nProNo)
    ElseIf sGameName = 市民大会 Then
        GetWinnerKey = CStr(nProNo) & sType
    Else
        ' 区分を取得
        If Trim(VLookupArea(nProNo, sMasterName, "種目区分")) = "" Then
            GetWinnerKey = CStr(nProNo) & sType
        Else
            GetWinnerKey = CStr(nProNo)
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
    
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' 優勝者範囲名
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)

    ' 削除
    Call DeleteWinnerSheet(sWinnerAreaName)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
        
    ' 出力する
    Call WriteWinnerList(sWinnerAreaName, sRecordAreaName, oWinnerList)

    ' 書式設定
    Call SetWinnerRecordStyle(sGameName)
    
    ' 印刷範囲の設定
    ActiveSheet.PageSetup.PrintArea = TableRangeAddress(sWinnerAreaName)

End Sub

'
' 優勝者書込み
'
' sWinnerAreaName   IN  優勝者範囲名
' sRecordAreaName   IN  大会記録範囲名
' oWinnerList       IN  優勝者リスト
'
Private Sub WriteWinnerList(sWinnerAreaName As String, sRecordAreaName As String, oWinnerList As Object)
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
        
    ' 大会記録毎
    Dim vKey As Range
    For Each vKey In GetAreaKeyData(sRecordAreaName)
        If oWinnerList.Exists(CStr(vKey.Value)) Then
            Set oWinners = oWinnerList.Item(CStr(vKey.Value))
            Dim vIdx As Variant
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteWinnerLine(sWinnerAreaName, sRecordAreaName, nRow, vKey, oWinner)
                nRow = nRow + 1
            Next vIdx
        End If
    Next vKey
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
' 優勝者シート記入
'
' sWinnerAreaName   IN  優勝者範囲名
' sRecordAreaName   IN  大会記録範囲名
' nRow              IN  優勝者の行数
' vKey              IN  大会記録の参照元の基準セル
' oWinner           IN  優勝者情報
'
Private Sub WriteWinnerLine(sWinnerAreaName As String, sRecordAreaName As String, nRow As Integer, vKey As Variant, oWinner As Object)
    
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "プロNo.")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "プロNo."))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "性別")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "性別"))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "距離")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "距離"))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "種目")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "種目"))
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "区分")) = GetOffset(vKey, GetAreaColumnIndex(sRecordAreaName, "区分"))
    
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "氏名")) = oWinner.Item("氏名")
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "所属")) = oWinner.Item("所属")
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "記録")) = oWinner.Item("記録")
    Cells(nRow, GetAreaColumnIndex(sWinnerAreaName, "大会新")) = oWinner.Item("大会新")

End Sub

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
    
    ' 大会記録の書込み
    Call WriteNewRecords(sGameName, oWinnerList)

    Call 各種設定名前定義(設定各種シート)

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
    
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
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
        
        ' プロNo＋区分の１位が未登録の場合
        If Not (oRecordList.Exists(sKey)) Then
            Set oWinners = CreateObject("Scripting.Dictionary")
            oWinners.Add oWinners.Count + 1, oWinner
            oRecordList.Add sKey, oWinners
        Else
            Set oWinners = oRecordList.Item(sKey)
            oWinners.Add oWinners.Count + 1, oWinner
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

    Dim sKey As String

    ' 優勝者毎
    Dim vCell As Range
    For Each vCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' 大会新なら格納する
        If GetOffset(vCell, GetAreaColumnIndex(sWinnerAreaName, "大会新")).Value = "大会新" Then
                
            ' キーを取得
            sKey = GetWinnerKey(sGameName, sMasterName, CInt(vCell.Value), _
                    GetOffset(vCell, GetAreaColumnIndex(sWinnerAreaName, "区分")))
                
            Set oWinner = CreateObject("Scripting.Dictionary")
            
            ' カラムの値の登録
            Dim vKey As Variant
            For Each vKey In GetRange(sWinnerAreaName).Rows(1).Columns()
                oWinner.Add STrimAll(vKey.Value), GetOffset(vCell, vKey.Column).Value
            Next vKey
            oWinner.Add GetAreaKeyName(sRecordAreaName), sKey
            oWinner.Add "年", nYear
                
            ' プロNo＋区分の１位が未登録の場合
            If Not (oRecordList.Exists(sKey)) Then
                Set oWinners = CreateObject("Scripting.Dictionary")
                oWinners.Add oWinners.Count + 1, oWinner
                oRecordList.Add sKey, oWinners
            Else
                Set oWinners = oRecordList.Item(sKey)
                ' 既に存在する場合はタイムを比較し古ければ削除
                For Each vKey In oWinners.Keys()
                    Set oWinnerOld = oWinners.Item(vKey)
                    If oWinner.Item("記録") < oWinnerOld.Item("記録") Then
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
    
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)
    
    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

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

    ' シートの保護
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = True
End Sub

'
' 大会記録書込み
'
' sGameName     IN  大会名
' oWinnerList   IN  優勝者リスト
'
Private Sub WriteNewRecordsOld(sGameName As String, oRecordList As Object)

    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Dim vVisible As Variant
    vVisible = SheetActivate(sSheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetProtect(False)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    ' 大会記録毎
    Dim vCell As Range
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        If oRecordList.Exists(vCell.Value) Then
            Set oWinners = oRecordList.Item(vCell.Value)
            Dim vIdx As Variant
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteRecordLine(sRecordAreaName, vCell, oWinner)
            Next vIdx
        End If
    Next vCell

    ' シートの保護
    Set oWorkSheet = SheetProtect(True, oWorkSheet)
    oWorkSheet.Visible = True
End Sub


'
' 大会記録シート記入
'
' sGameName     IN  大会名
' vCell         IN  参照元の基準セル
' oWinner       IN  優勝者情報
'
Private Sub WriteRecordLine(sAreaName As String, vCell As Variant, oWinner As Object)

    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "氏名") - 1).Value = oWinner.Item("氏名")
    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "所属") - 1).Value = oWinner.Item("所属")
    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "記録") - 1).Value = oWinner.Item("記録")
    vCell.Offset(0, GetAreaColumnIndex(sAreaName, "年") - 1).Value = oWinner.Item("年")

End Sub
