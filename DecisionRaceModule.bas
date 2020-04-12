Attribute VB_Name = "DecisionRaceModule"
'
' 組み合わせ決定
'
' レースNoと組、レーンを設定する
'
Sub 組み合わせ決定()
    ' イベント発生を抑制
    Call EventChange(False)

    ' エクセルシートを選択
    Call SheetActivate(sEntrySheetName)

    ' 出力用ワークブック
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' 出力用ワークシート
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' ProNo、ソート区分、申込み時間でソート
    Call SortByProNo(oWorkSheet, sEntryTableName)
    
    ' 組み合わせ作成
    Call SetHeatLaneOrder(oWorkSheet, sEntryTableName)
    
    ' レースNo, レーンでソート
    Call SortByRace(oWorkSheet, sEntryTableName)
    
    ' イベント発生を発生
    Call EventChange(True)

    ' シートを保存
    oWorkBook.Save
End Sub

'
' 組み合わせ作成
'
' エントリー一覧を読込み、組み合わせを作成する
' エントリー一覧はProNo、ソート区分、申込み時間でソートされている前提
'
' oWorkSheet    IN      ワークシート
' sTableName    IN      テーブル名
'
Sub SetHeatLaneOrder(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate
    
    ' エントリー一覧
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' データを格納
    Call ReadProNo(sTableName, oEntryList)

    Dim nRaceNo As Integer      ' レースNo
    Dim nCount As Integer       ' プロNo毎の件数
    Dim nRow As Integer         ' カレントの行番号
    Dim nHeat As Integer        ' 組番号
    Dim nPreHeat As Integer     ' これまでの組番号
    Dim nNum As Integer
    nTotalHeat = 0
    ' プログラムNo毎
    For Each nProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(nProNo)
        nCount = oProNo.Count
        nPreHeat = 0
        
        ' 組毎
        For Each nOrder In oProNo.Keys
            ' カレント行番号
            nRow = oProNo.Item(nOrder)
            ' 組番号を決定
            nHeat = GetOrderHeat(nCount, Int(nOrder))
            If nPreHeat <> nHeat Then
                ' 組番号が変わった場合
                nNum = 1
                nPreHeat = nHeat
                nRaceNo = nRaceNo + 1   ' レースNoをインクリメント
            Else
                nNum = nNum + 1
            End If
            ' レースNo、組の書込み
            Cells(nRow, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
            Cells(nRow, Range(sTableName & "[組]").Column).Value = nHeat
                      
            ' 横須賀選手権水泳大会
            If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
                Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
            ' 横須賀市民体育大会
            ElseIf GetRange("大会名").Value = "横須賀市民体育大会" Then
                Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
            Else
                ' 横須賀マスターズ
                If Cells(nRow, Range(sTableName & "[ソート区分]").Column).Value <> "" Then
                    Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane2(nHeat, oProNo.Count, nNum, False)
                ' 学童
                Else
                    Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane(nHeat, oProNo.Count, nNum, False)
                End If
            End If
        Next
    Next

End Sub

'
' プロNoをキーにデータ格納
'
' プロNo毎に通番を振ってその行番号を格納する
'
' sTableName    IN      テーブル名
' oEntryList    OUT     エントリー番号
'　└プロNo
'　　└通番：行番号
'
Sub ReadProNo(sTableName As String, oEntryList As Object)
    ' データを格納
    Dim oProNo As Object
    For Each cProNo In Range(sTableName & "[プロNo]")
        If Not oEntryList.Exists(cProNo.Value) Then
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add cProNo.Value, oProNo
        End If
        oProNo.Add oProNo.Count + 1, cProNo.Row
    Next
End Sub

'
' 組番号決定
'
' プロNoの総人数と順番から組番号を算出する
'
' 組数、１組目の人数、２組目の人数を算出することで
' 組番号が算出可能
'
' nTotalNum     IN      プロNoの総人数
' nOrder        IN      順番
'
Function GetOrderHeat(nTotalNum As Integer, nOrder As Integer)

    ' 組数を算出
    Dim nHeats As Integer
    nHeats = GetHeats(nTotalNum)
    
    ' １組目人数算出
    Dim nFirstHeatNumber As Integer
    nFirstHeatNumber = GetFirstHeatNumber(nTotalNum)
    
    ' ２組目人数算出
    Dim nSecondHeatNumber As Integer
    nSecondHeatNumber = GetSecondHeatNumber(nTotalNum)
    
    ' １組目の場合
    If nOrder <= nFirstHeatNumber Then
        GetOrderHeat = 1
    ' ２組目の場合
    ElseIf nOrder <= nFirstHeatNumber + nSecondHeatNumber Then
        GetOrderHeat = 2
    ' ３組目以降の場合
    Else
        GetOrderHeat = GetHeats(nOrder - (nFirstHeatNumber + nSecondHeatNumber)) + 2
    End If

End Function
'
' 組数算出
'
' 組数はレースの総人数÷１レースの人数
'
' nTotalNum     IN      レースの総人数
'
Function GetHeats(nTotalNum As Integer)

    GetHeats = Application.WorksheetFunction.RoundUp(nTotalNum / nNumberOfRace, 0)

End Function

'
' １組目人数算出
'
' 総人数が３名以上いる場合、最低３名は１組目に残す
'
' nTotalNum     IN      レースの総人数
'
Function GetFirstHeatNumber(nTotalNum As Integer)

    If nTotalNum <= nNumberOfRace Then
        GetFirstHeatNumber = nTotalNum
    ElseIf nTotalNum Mod nNumberOfRace = 0 Then
        GetFirstHeatNumber = nNumberOfRace
    ElseIf nTotalNum Mod nNumberOfRace <= nMinLaneOfRace Then
        GetFirstHeatNumber = nMinLaneOfRace
    Else
        GetFirstHeatNumber = nTotalNum Mod nNumberOfRace
    End If

End Function

'
' ２組目人数算出
'
' １組目に回す人数によって２組目も変化する
'
' nTotalNum     IN      レースの総人数
'
Function GetSecondHeatNumber(nTotalNum As Integer)

    If nTotalNum <= nNumberOfRace Then
        GetSecondHeatNumber = 0
    ElseIf nTotalNum Mod nNumberOfRace = 0 Then
        GetSecondHeatNumber = nNumberOfRace
    ElseIf nTotalNum Mod nNumberOfRace <= nMinLaneOfRace Then
        GetSecondHeatNumber = nNumberOfRace + (nTotalNum Mod nNumberOfRace - nMinLaneOfRace)
    Else
        GetSecondHeatNumber = nNumberOfRace
    End If

End Function

'
' レーン決定（学童用）
'
' レーンは競技規則の単純方式で並べる
'
' nHeat         IN      組番号
' nTotalNum     IN      プロNoの総人数
' nOrder        IN      順番
' bFlag         In      True：通常／False：逆順
'
Function GetLane(nHeat As Integer, nTotalNum As Integer, nOrder As Integer, Optional bFlag As Boolean = True)

    Dim nMax As Integer
    Dim nNum As Integer
    
    If nHeat = 1 Then
        nMax = GetFirstHeatNumber(nTotalNum)
    ElseIf nHeat = 2 Then
        nMax = GetSecondHeatNumber(nTotalNum)
    Else
        nMax = nNumberOfRace
    End If
    
    nNum = nMax - nOrder + 1
    
    If bFlag Then
        ' 4->5->3->6->2->7->1
        GetLane = nCenterLane - Application.WorksheetFunction.Power(-1, nNum - 1) _
                    * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        ' 4->3->5->2->6->1->7(なぜか学童はこちらだった)
        GetLane = nCenterLane + Application.WorksheetFunction.Power(-1, nNum - 1) _
                    * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    End If

End Function

'
' レーン決定（マスターズ用）
'
' レーンは昇順で並べる
'
' nHeat         IN      組番号
' nTotalNum     IN      プロNoの総人数
' nOrder        IN      順番
' bFlag         In      True：通常／False：逆順
'
Function GetLane2(nHeat As Integer, nTotalNum As Integer, nOrder As Integer, Optional bFlag As Boolean = True)

    Dim nMax As Integer
    
    If nHeat = 1 Then
        nMax = GetFirstHeatNumber(nTotalNum)
    ElseIf nHeat = 2 Then
        nMax = GetSecondHeatNumber(nTotalNum)
    Else
        nMax = nNumberOfRace
    End If
    
    If bFlag Then
        ' 4->5->3->6->2->7->1
        GetLane2 = nCenterLane + nOrder - Application.WorksheetFunction.RoundUp(nMax / 2, 0)
    Else
        ' 4->3->5->2->6->1->7(なぜか学童はこちらだった)
        GetLane2 = nCenterLane + Application.WorksheetFunction.RoundUp(nMax / 2, 0) - (nMax - nOrder) - 1
    End If
End Function


'
' レース番号修正
'
'
Sub レース番号修正()
    ' イベント発生を抑制
    Call EventChange(False)

    ' 出力用ワークブック
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook

    ' 出力用シート
    Call SheetActivate(sEntrySheetName)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' 再ソート
    Call SortByRace(oWorkSheet, sEntryTableName)
    
    ' レース番号修正
    Call SetRaceNo(oWorkSheet, sEntryTableName)

    ' イベント発生を再開
    Call EventChange(True)
End Sub

'
' レース番号修正
'
' oWorkSheet    IN  ワークシート
' sTableName    IN  テーブル名
'
Sub SetRaceNo(oWorkSheet As Worksheet, sTableName As String)
    
    ' ProNo、組の重複チェック
    Dim oEntryList As Object
    Call ReadEntrySheet(sTableName, oEntryList)
    
    ' エントリー一覧
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' データを格納＆レース番号のチェック
    Call ReadRace(sTableName, oEntryList)
    
    Dim nRaceNo As Integer
    nRaceNo = 1
    
    Dim nCurrentRaceNo As Integer
    nCurrentRaceNo = 0

    ' レース番号修正
    Dim nNewNo As Integer
    nNewNo = 1
    Dim nRow As Integer
    For Each vRaceNo In oEntryList.Keys
        Set oRaceNo = oEntryList.Item(vRaceNo)
        
        ' レーン
        For Each vLane In oRaceNo.Keys
            nRow = oRaceNo.Item(vLane)
            
            ' レース番号
            Cells(nRow, Range(sTableName & "[レースNo]").Column).Value = nNewNo
        Next
        nNewNo = nNewNo + 1
    Next

End Sub

'
' レース番号の読込み
'
' 読込みながらレース番号中のレーン重複チェックも行う
'
' sTableName    IN      テーブル名
' oEntryList    I/O     エントリー一覧
'
Sub ReadRace(sTableName As String, oEntryList As Object)
    Dim nLane As Integer
    Dim oRaceNo As Object
    For Each cRaceNo In Range(sTableName & "[レースNo]")
        ' 存在しないレースNoの場合は
        If Not oEntryList.Exists(cRaceNo.Value) Then
            ' エントリー一覧に登録する
            Set oRaceNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add cRaceNo.Value, oRaceNo
        End If
        ' レーン番号を取得
        nLane = cRaceNo.Offset(0, Range(sTableName & "[レーン]").Column - Range(sTableName & "[レースNo]").Column).Value
        ' レース番号に対するレーンの重複チェック
        If oRaceNo.Exists(nLane) Then
            MsgBox "レースNo：" & Str(cRaceNo.Value) & vbCrLf & _
                    "レーン　：" & Str(nLane) & vbCrLf & _
                    "が重複しています。"
            Range(sTableName).Parent.Activate
            Range(Cells(cRaceNo.Row, Range(sTableName & "[レースNo]").Column), _
                    Cells(cRaceNo.Row, Range(sTableName & "[レーン]").Column)).Select
            cRaceNo.Activate
            End
        Else
            oRaceNo.Add nLane, cRaceNo.Row
        End If
    Next cRaceNo
End Sub

'
' レースNo、組でソートする
'
' oWorkSheet    IN      ワークシート
' sTableName    IN      テーブル名
'
Sub SortByRace(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate

    With ActiveSheet.ListObjects(sTableName).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range(sTableName & "[レースNo]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range(sTableName & "[レーン]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

