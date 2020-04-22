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
    
    Dim sProType As String
    Dim sKey As String
    Dim sName As String
    Dim sTeam As String
    Dim sType As String
    Dim nTime As Long
    Dim nRecord As Variant
    
    ' プログラム番号毎
    For Each nProNo In GetAreaKeyData(sMasterName)
        sProType = Trim(VLookupArea(nProNo, sMasterName, "区分"))
        
        ' プログラム番号内から１位を探す
        For Each oCell In GetRange("プログラム番号" & CStr(nProNo))
            ' １位の場合
            If oCell.Offset(0, Range("Header順位").Column - Range("HeaderプロNo").Column).Value = 1 Then
                
                Set oWinner = CreateObject("Scripting.Dictionary")
                
                sName = oCell.Offset(0, Range("Header氏名").Column - Range("HeaderプロNo").Column).Value
                sTeam = oCell.Offset(0, Range("Header所属").Column - Range("HeaderプロNo").Column).Value
                If sProType = "" Then
                    sType = oCell.Offset(0, Range("Header区分").Column - Range("HeaderプロNo").Column).Value
                Else
                    sType = ""
                End If
                nTime = oCell.Offset(0, Range("Header時間").Column - Range("HeaderプロNo").Column).Value
                nRecord = oCell.Offset(0, Range("Header大会記録").Column - Range("HeaderプロNo").Column).Value
                sKey = CStr(nProNo) & sType
                
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
    Next nProNo
End Sub


'
' 優勝者書込み
'
' sGameName     IN  大会名
' oWinnerList   IN  優勝者リスト
'
Sub WriteWinner(sGameName As String, oWinnerList As Object)

    ' 優勝者シート
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
        If oWinnerList.Exists(vKey.Value) Then
            Set oWinners = oWinnerList.Item(vKey.Value)
            For Each vIdx In oWinners
                Set oWinner = oWinners.Item(vIdx)
                Call WriteWinnerLine(sWinnerAreaName, nRow, vKey, oWinner)
                nRow = nRow + 1
            Next vIdx
        End If
    Next vKey

    ' 書式設定
    Call SetWinnerRecordStyle(sGameName, sWinnerAreaName, sRecordAreaName)
    
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
' sGameName     IN  大会名
' nRow          IN  行数
' vKey          IN  参照元の基準セル
' oWinner       IN  優勝者情報
'
Sub WriteWinnerLine(sAreaName As String, nRow As Integer, vKey As Variant, oWinner As Object)

    Dim nProNoCol As Integer
    nProNoCol = GetAreaColumnIndex(sAreaName, "プロNo.")
    Dim nGenDist As Integer
    nGenDist = GetAreaColumnIndex(sAreaName, "種")
    Dim nStyle As Integer
    nStyle = GetAreaColumnIndex(sAreaName, "目")
    Dim nTypeCol As Integer
    nTypeCol = GetAreaColumnIndex(sAreaName, "区分")
    Dim nNameCol As Integer
    nNameCol = GetAreaColumnIndex(sAreaName, "氏名")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sAreaName, "所属")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sAreaName, "記録")
    Dim nRecordCol As Integer
    nRecordCol = GetAreaColumnIndex(sAreaName, "大会新")
    
    Cells(nRow, nProNoCol) = vKey.Offset(0, nProNoCol)
    Cells(nRow, nGenDist) = vKey.Offset(0, nGenDist)
    Cells(nRow, nStyle) = vKey.Offset(0, nStyle)
    Cells(nRow, nTypeCol) = vKey.Offset(0, nTypeCol)
    
    Cells(nRow, nNameCol) = oWinner.Item("氏名")
    Cells(nRow, nTeamCol) = oWinner.Item("所属")
    Cells(nRow, nTimeCol) = oWinner.Item("記録")
    Cells(nRow, nRecordCol) = oWinner.Item("大会新")

End Sub

'
' 優勝者シートの書式を大会記録からコピーする
'
' sGameName         IN  大会名
' sWinnerAreaName   IN  優勝者の範囲名
' sRecordAreaName   IN  大会記録の範囲名
'
Sub SetWinnerRecordStyle(sGameName As String, sWinnerAreaName As String, sRecordAreaName As String)

    ' 大会記録シート
    Dim sSheetName As String
    sSheetName = GetRecordSheetName(sGameName)
    Sheets(sSheetName).Select
    GetRange(sRecordAreaName).Offset(1, 1).Resize(1).Copy
    
    ' 優勝者シート
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select
    Dim oRange As Range
    Set oRange = TableRange(sWinnerAreaName)
    oRange.Offset(1, 0).Resize(oRange.Rows.Count - 1, oRange.Columns.Count).Select
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
    Call ReadNewRecords(sGameName, oWinnerList)
    
    ' 大会記録の書込み
    Call WriteNewRecords(sGameName, oWinnerList)

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
Sub ReadNewRecords(sGameName As String, oRecordList As Object)
    
    ' 優勝者シート
    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select

    ' 優勝者の範囲
    Dim sWinnerAreaName As String
    sWinnerAreaName = GetWinnerAreaName(sGameName)
    
    ' 大会年
    Dim nYear As Integer
    nYear = GetRange("大会年").Value
    
    Dim oWinners As Object
    Dim oWinner As Object

    Dim nProNoCol As Integer
    nProNoCol = GetAreaColumnIndex(sWinnerAreaName, "プロNo.")
    Dim nTypeCol As Integer
    nTypeCol = GetAreaColumnIndex(sWinnerAreaName, "区分")
    Dim nNameCol As Integer
    nNameCol = GetAreaColumnIndex(sWinnerAreaName, "氏名")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sWinnerAreaName, "所属")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sWinnerAreaName, "記録")
    Dim nRecordCol As Integer
    nRecordCol = GetAreaColumnIndex(sWinnerAreaName, "大会新")
    
    ' 優勝者毎
    For Each oCell In RowRange(GetRange(sWinnerAreaName).Columns(1).Address).Offset(1)
        
        ' 大会新なら格納する
        If oCell.Offset(0, nRecordCol - nProNoCol).Value = "大会新" Then
                
            Set oWinner = CreateObject("Scripting.Dictionary")
                
            sType = oCell.Offset(0, nTypeCol - nProNoCol).Value
            sName = oCell.Offset(0, nNameCol - nProNoCol).Value
            sTeam = oCell.Offset(0, nTeamCol - nProNoCol).Value
            nTime = oCell.Offset(0, nTimeCol - nProNoCol).Value
            
            ' 学童マスターズで小学が含まれる場合
            If sType Like "小学*" Then
                sKey = CStr(oCell.Value)
            Else
                sKey = CStr(oCell.Value) & sType
            End If
                
            oWinner.Add "氏名", sName
            oWinner.Add "所属", sTeam
            oWinner.Add "記録", nTime
            oWinner.Add "年", nYear
                
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
    Next oCell
End Sub


'
' 優勝者書込み
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
