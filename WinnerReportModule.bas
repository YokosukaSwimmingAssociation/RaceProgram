Attribute VB_Name = "WinnerReportModule"
Sub 優勝者一覧作成()

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
Sub ReadWinner(sGameName As String, oWinnerList As Object)
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)
    
    ' １位リスト
    Dim oWinners As Object
    Set oWinners = Nothing
    ' １位情報
    Dim oWinner As Object
    Set oWinner = Nothing
    
    Dim sKey As String
    Dim sName As String
    Dim sTeam As String
    Dim nTime As Long
    Dim nRecord As Variant
    Dim bFlag As Boolean
    
    ' プログラム番号毎
    For Each nProNo In GetAreaKeyData(sMasterName)
        ' 選手権の予選は飛ばす
        If sGameName = 選手権大会 Then
             If VLookupArea(nProNo, "選手権種目区分", "予選／決勝") = "予選" Then
                bFlag = False
             Else
                bFlag = True
             End If
        Else
            bFlag = True
        End If
        
        ' 決勝（タイム決勝）の場合
        If bFlag Then
            ' プログラム番号内から１位を探す
            For Each oCell In GetRange("プログラム番号" & CStr(nProNo))
                ' １位の場合
                If oCell.Offset(0, Range("Header順位").Column - Range("HeaderプロNo").Column).Value = 1 Then
                    
                    Set oWinner = CreateObject("Scripting.Dictionary")
                    
                    sName = oCell.Offset(0, Range("Header氏名").Column - Range("HeaderプロNo").Column).Value
                    sTeam = oCell.Offset(0, Range("Header所属").Column - Range("HeaderプロNo").Column).Value
                    nTime = oCell.Offset(0, Range("Header時間").Column - Range("HeaderプロNo").Column).Value
                    nRecord = oCell.Offset(0, Range("Header大会記録").Column - Range("HeaderプロNo").Column).Value
                    sKey = GetWinnerKey(sGameName, sMasterName, CInt(nProNo), _
                        oCell.Offset(0, Range("Header区分").Column - Range("HeaderプロNo").Column).Value)
    
                    oWinner.Add "氏名", sName
                    oWinner.Add "所属", sTeam
                    oWinner.Add "記録", nTime
                    If Not IsNumeric(nRecord) Or nTime <= nRecord Then
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
                End If
            Next oCell
        End If
    Next nProNo
End Sub

'
' 大会毎の大会記録の区分
'
' sGameName     IN  大会名
' sMasterName   IN  マスター名
' nProNo        IN  種目番号
' sType         IN  区分
'
Function GetWinnerKey(sGameName As String, sMasterName As String, nProNo As Integer, sType As String)

    If sGameName = 選手権大会 Then
        GetWinnerKey = CStr(nProNo)
    ElseIf sGameName = 市民大会 Then
        GetWinnerKey = CStr(nProNo) & sType
    Else
        ' 区分を取得
        If Trim(VLookupArea(nProNo, sMasterName, "区分")) = "" Then
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
Sub WriteWinner(sGameName As String, oWinnerList As Object)

    ' 優勝者シートを選択し保護を解除
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' 優勝者範囲名
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)

    ' 削除
    Call DeleteWinnerSheet(sWinnerAreaName)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
        
    ' 大会記録毎
    For Each vKey In GetAreaKeyData(sRecordAreaName)
        If oWinnerList.Exists(CStr(vKey.Value)) Then
            Set oWinners = oWinnerList.Item(CStr(vKey.Value))
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteWinnerLine(sWinnerAreaName, sRecordAreaName, nRow, vKey, oWinner)
                nRow = nRow + 1
            Next vIdx
        End If
    Next vKey

    ' 書式設定
    Call SetWinnerRecordStyle(sGameName)
    
    ' 印刷範囲の設定
    ActiveSheet.PageSetup.PrintArea = TableRangeAddress(sWinnerAreaName)

End Sub

Sub DeleteWinnerSheet(sWinnerAreaName As String)
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    If Cells(oRange.Row + 1, oRange.Column) <> "" Then
        oRange.Offset(1, 0).Resize(oRange.Rows().Count - 1).EntireRow.Delete
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
Sub WriteWinnerLine(sWinnerAreaName As String, sRecordAreaName As String, nRow As Integer, vKey As Variant, oWinner As Object)
    
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
Sub SetWinnerRecordStyle(sGameName As String)

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
' 大会記録シートの書式を優勝者からコピーする
'
' sGameName         IN  大会名
'
Sub SetRecordWinnerStyle(sGameName As String)

    ' 優勝者シート
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    
    ' 優勝者範囲名
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    GetRange(sWinnerAreaName).Offset(1, 0).Resize(1).Copy
    
    ' 大会記録シート
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    
     ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)
    
    Dim oRange As Range
    Set oRange = TableRange(sRecordAreaName)
    If Cells(oRange.Row + 1, oRange.Column) <> "" Then
        oRange.Offset(1, 1).Resize(oRange.Rows.Count - 1, oRange.Columns.Count - 1).Select
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
Function GetWinnerSheetName(sGameName As String)
    If sGameName = "横須賀選手権水泳大会" Then
        GetWinnerSheetName = "選手権大会優勝者"
    ElseIf sGameName = "横須賀市民体育大会" Then
        GetWinnerSheetName = "市民大会優勝者"
    Else
        GetWinnerSheetName = "学童マスターズ優勝者"
    End If
End Function

'
' 優勝者範囲名
'
' sGameName     IN      大会名
'
Function GetWinnerAreaName(sGameName As String)
    If sGameName = "横須賀選手権水泳大会" Then
        GetWinnerAreaName = "選手権大会優勝者"
    ElseIf sGameName = "横須賀市民体育大会" Then
        GetWinnerAreaName = "市民大会優勝者"
    Else
        GetWinnerAreaName = "学マ大会優勝者"
    End If
End Function

'
' 大会記録シート名
'
' sGameName     IN      大会名
'
Function GetRecordSheetName(sGameName As String)
    If sGameName = "横須賀選手権水泳大会" Then
        GetRecordSheetName = "選手権大会記録"
    ElseIf sGameName = "横須賀市民体育大会" Then
        GetRecordSheetName = "市民大会記録"
    Else
        GetRecordSheetName = "学童マスターズ大会記録"
    End If
End Function


'
' 大会記録範囲名
'
' sGameName     IN      大会名
'
Function GetRecordAreaName(sGameName As String)
    If sGameName = "横須賀選手権水泳大会" Then
        GetRecordAreaName = "選手権大会記録"
    ElseIf sGameName = "横須賀市民体育大会" Then
        GetRecordAreaName = "市民大会記録"
    Else
        GetRecordAreaName = "学マ大会記録"
    End If
End Function

Sub 大会記録更新()

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

    Call ワークブック名前定義

    ' イベント発生を発生
    Call EventChange(True)

    ' 保存
    ActiveWorkbook.Save

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
Sub ReadRecords(sGameName As String, oRecordList As Object)
    
    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim sKey As String
        
    ' 大会記録毎
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        sKey = vCell.Value
        
        Set oWinner = CreateObject("Scripting.Dictionary")
        
        ' カラムの値の登録
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
Sub ReadNewRecords(sGameName As String, oRecordList As Object)
    
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

    ' 優勝者毎
    For Each oCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' 大会新なら格納する
        If GetOffset(oCell, GetAreaColumnIndex(sWinnerAreaName, "大会新")).Value = "大会新" Then
                
            ' キーを取得
            sKey = GetWinnerKey(sGameName, sMasterName, CInt(oCell.Value), _
                    GetOffset(oCell, GetAreaColumnIndex(sWinnerAreaName, "区分")))
                
            Set oWinner = CreateObject("Scripting.Dictionary")
            
            ' カラムの値の登録
            For Each vKey In GetRange(sWinnerAreaName).Rows(1).Columns()
                oWinner.Add STrimAll(vKey.Value), GetOffset(oCell, vKey.Column).Value
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
    Next oCell
End Sub

'
' 大会記録書込み
'
' sGameName     IN  大会名
' oWinnerList   IN  優勝者リスト
'
Sub WriteNewRecords(sGameName As String, oRecordList As Object)

    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    ' 削除
    Call DeleteWinnerSheet(sRecordAreaName)

    Dim oWinners As Object
    Dim oWinner As Object
        
    Dim nRow As Integer
    nRow = 2
    
    ' 大会記録毎
    For Each vKey In oRecordList.Keys()
        Set oWinners = oRecordList.Item(vKey)
        For Each vIdx In oWinners
            Set oWinner = oWinners.Item(vIdx)
            
            ' カラムの値の書込み
            For Each vCell In GetRange(sRecordAreaName).Rows(1).Columns()
                Cells(nRow, vCell.Column) = oWinner.Item(STrimAll(vCell.Value))
            Next vCell
            
            nRow = nRow + 1
        Next vIdx
    Next vKey

    ' 書式設定
    Call SetRecordWinnerStyle(sGameName)

    ' シートの保護
    Call SheetProtect(True)
End Sub

'
' 大会記録書込み
'
' sGameName     IN  大会名
' oWinnerList   IN  優勝者リスト
'
Sub WriteNewRecordsOld(sGameName As String, oRecordList As Object)

    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    Call SheetProtect(False)

    ' 大会記録範囲名
    Dim sRecordAreaName As String
    sRecordAreaName = GetRecordAreaName(sGameName)

    Dim oRange As Range
    Set oRange = GetRange(sRecordAreaName)
        
    Dim oWinners As Object
    Dim oWinner As Object
        
    ' 大会記録毎
    For Each vCell In GetAreaKeyData(sRecordAreaName)
        If oRecordList.Exists(vCell.Value) Then
            Set oWinners = oRecordList.Item(vCell.Value)
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteRecordLine(sRecordAreaName, vCell, oWinner)
            Next vIdx
        End If
    Next vCell

    ' シートの保護
    Call SheetProtect(True)
End Sub


'
' 大会記録シート記入
'
' sGameName     IN  大会名
' vCell         IN  参照元の基準セル
' oWinner       IN  優勝者情報
'
Sub WriteRecordLine(sAreaName As String, vCell As Variant, oWinner As Object)

    nNameCol = GetAreaColumnIndex(sAreaName, "氏名")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sAreaName, "所属")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sAreaName, "記録")
    Dim nRecordCol As Integer
    nRecordCol = GetAreaColumnIndex(sAreaName, "年")
    
    vCell.Offset(0, nNameCol - 1).Value = oWinner.Item("氏名")
    vCell.Offset(0, nTeamCol - 1).Value = oWinner.Item("所属")
    vCell.Offset(0, nTimeCol - 1).Value = oWinner.Item("記録")
    vCell.Offset(0, nRecordCol - 1).Value = oWinner.Item("年")

End Sub
