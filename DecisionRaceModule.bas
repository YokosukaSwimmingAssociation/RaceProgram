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
    Call SheetActivate(S_ENTRY_SHEET_NAME)

    ' 出力用ワークブック
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' 出力用ワークシート
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' ProNo、ソート区分、申込み時間でソート
    Call SortByProNo(oWorkSheet, S_ENTRY_TABLE_NAME)
    
    ' 組み合わせ作成
    Call SetHeatLaneOrder(oWorkSheet, S_ENTRY_TABLE_NAME)
    
    ' レースNo, レーンでソート
    Call SortByRace(oWorkSheet, S_ENTRY_TABLE_NAME)
    
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
    
    ' 組み合わせが学マだけ逆な対応（検証用の暫定）
    Dim bFlag As Boolean
    If GetRange("大会名").Value = "学童マスターズ大会" Then
        bFlag = False
    Else
        bFlag = True
    End If
    
    ' 組の最小人数
    Dim nMinNumberOfRace As Integer
    If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
        nMinNumberOfRace = N_MIN_NUMBER_OF_RACE
    Else
        nMinNumberOfRace = N_MIN_NUMBER_OF_RACE2
    End If
    
    
    ' エントリー一覧
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    Dim oProNo As Object
    
    ' データを格納
    Call ReadProNo(sTableName, oEntryList)

    Dim nRow As Integer             ' カレントの行番号
    
    Dim nNumOfProNo As Integer      ' プログラムNo毎の人数
    Dim nRaceNo As Integer          ' レースNo
    Dim nHeat As Integer            ' 組番号
    Dim nHeats As Integer           ' 組数
    Dim nNumOfHeat As Integer       ' 組の人数
    Dim nNumOfHeats() As Integer    ' 組毎の人数
    Dim nNumOfSortType As Integer   ' ソート区分毎の残り人数
    Dim nRemNumber As Integer       ' 組の残り人数
    
    Dim nOrder As Integer           ' ProNo毎の順序（逆順）
    Dim nStartLane As Integer       ' 組の開始位置
    Dim nCenterLane As Integer      ' センターレーン
    Dim nMax As Integer             ' 組の中でLane決定する人数
    Dim nNum As Integer             ' Lene決定する人数の中の順位
    Dim bAverage As Boolean         ' 平均分け方式の利用有無
    Dim nMaxNum As Integer          ' 平均分け方式の総人数
    
    nRaceNo = 0
    
    ' プログラムNo毎
    For Each nProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(nProNo)
        
        ' プログラムNo毎の人数
        nNumOfProNo = GetNumberOfProNo(oProNo)
        
        ' 組数を算出
        nHeats = GetHeats(nNumOfProNo)
        Call GetNumberOfHeat(nNumOfProNo, nHeats, nNumOfHeats, nMinNumberOfRace)
        
        ' 平均分け方式を利用するケース
        bAverage = False
        If GetRange("大会名").Value = "横須賀選手権水泳大会" And nHeats >= N_AVERAGE_DEC_RACE Then
            If VLookupArea(nProNo, "選手権種目区分", "予選／決勝") = "予選" Then
                bAverage = True
            End If
        End If
        
        ' ProNo毎の選手位置
        nOrder = 1
        
        ' 組毎
        For nHeat = 1 To nHeats
            ' レースNoをインクリメント
            nRaceNo = nRaceNo + 10
            
            ' 組の人数
            nNumOfHeat = nNumOfHeats(nHeat - 1)
            ' 組の残り人数
            nRemNumber = nNumOfHeat
            
            ' 平均分け方式
            If bAverage And nHeat = (nHeats - N_AVERAGE_DEC_RACE) + 1 Then
                nMaxNum = 0
                For i = nHeat To nHeat + N_AVERAGE_DEC_RACE - 1
                    nMaxNum = nMaxNum + nNumOfHeats(i - 1)
                Next i
                Call AverageMethod(nRaceNo, CInt(nHeat), nMaxNum, nOrder, oProNo, sTableName)
                Exit For
            End If
        
            ' 組の開始位置
            nStartLane = GetStartLane(nNumOfHeat, GetCenterLane(N_NUMBER_OF_RACE, N_MIN_LANE_OF_RACE), bFlag)
            
            ' 組の人数が残っている間
            While nRemNumber > 0
                ' ソート区分毎の残り人数
                nNumOfSortType = GetNumberOfSortType(nOrder, oProNo)
                If nNumOfSortType <= nRemNumber Then
                    ' Lane決定する人数
                    nMax = nNumOfSortType
                Else
                    nMax = nRemNumber
                End If
                
                ' レーンを決定する
                For nNum = 1 To nMax
                    ' カレント行番号
                    nRow = GetProNoRow(nOrder, oProNo)
                
                    ' レースNo、組の書込み
                    Cells(nRow, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
                    Cells(nRow, Range(sTableName & "[組]").Column).Value = nHeat
                
                    ' レースNo、組、レーンを記述
                    nCenterLane = GetCenterLane(nMax, nStartLane, bFlag)
                    Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane(nCenterLane, nMax, nNum, bFlag)
                
                    ' 順番をインクリメント
                    nOrder = nOrder + 1
                Next
            
                ' 開始位置を変更
                nStartLane = nStartLane + nMax
            
                ' 残り人数を減算
                nRemNumber = nRemNumber - nMax
            Wend
        Next nHeat
    
    Next nProNo

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
        Set oRow = CreateObject("Scripting.Dictionary")
        oRow.Add "Row", cProNo.Row
        oRow.Add "SortType", Trim(Cells(cProNo.Row, Range(sTableName & "[ソート区分]").Column).Value)
        oProNo.Add oProNo.Count + 1, oRow
    Next
End Sub

'
' プロNo、ソート区分の件数
'
' oProNo        IN      プロNo毎のエントリー配列
'
Function GetNumberOfProNo(oProNo As Object)
    GetNumberOfProNo = oProNo.Count
End Function

'
' プロNoの順番の行
'
' nIndex        IN      順序番号
' oProNo        IN      プロNo毎のエントリー配列
'
Function GetProNoRow(nIndex As Integer, oProNo As Object)
    Dim oRow As Object
    Set oRow = oProNo.Item(nIndex)
    GetProNoRow = oRow.Item("Row")
End Function

'
' 開始位置から同一ソート区分の件数
'
' nIndex        IN      開始位置
' oProNo        IN      プロNo毎のエントリー配列
'
Function GetNumberOfSortType(nIndex As Integer, oProNo As Object)
    
    GetNumberOfSortType = 1
    Dim oRow As Object
    Set oRow = oProNo.Item(nIndex)
    Dim sSortType As String
    sSortType = oRow.Item("SortType")
    
    For i = nIndex + 1 To oProNo.Count
        Set oRow = oProNo.Item(i)
        If oRow.Item("SortType") = sSortType Then
            GetNumberOfSortType = GetNumberOfSortType + 1
        Else
            Exit Function
        End If
    Next
End Function


'
' 組人数配列を設定
'
' nTotalNum     IN      プロNoのエントリー数
' nHeat         IN      組数
' nNumberOfHeat() OUT   組毎の人数配列
' nMinNumberOfRace IN   組の最小人数
'
Sub GetNumberOfHeat(nTotalNum As Integer, nHeats As Integer, nNumberOfHeat() As Integer, nMinNumberOfRace As Integer)
    
    ReDim nNumberOfHeat(nHeats - 1) As Integer
    
    ' １組目人数算出
    nNumberOfHeat(0) = GetFirstHeatNumber(nTotalNum, nMinNumberOfRace)
    
    ' ２組目人数算出
    If nHeats >= 2 Then
        nNumberOfHeat(1) = GetSecondHeatNumber(nTotalNum, nMinNumberOfRace)
    
        ' ３組目以降
        If nHeats > 2 Then
            For i = 2 To nHeats - 1
                nNumberOfHeat(i) = N_NUMBER_OF_RACE
            Next
        End If
    End If
End Sub

'
' 組数算出
'
' 組数はレースの総人数÷１レースの人数
'
' nTotalNum     IN      レースの総人数
'
Function GetHeats(nTotalNum As Integer)

    GetHeats = Application.WorksheetFunction.RoundUp(nTotalNum / N_NUMBER_OF_RACE, 0)

End Function

'
' １組目人数算出
'
' 総人数が最小人数以上いる場合、最小人数は１組目に残す
'
' nTotalNum     IN      レースの総人数
' nMinNumberOfRace IN   組の最小人数
'
Function GetFirstHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer)

    If nTotalNum <= N_NUMBER_OF_RACE Then
        GetFirstHeatNumber = nTotalNum
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE = 0 Then
        GetFirstHeatNumber = N_NUMBER_OF_RACE
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE <= nMinNumberOfRace Then
        GetFirstHeatNumber = nMinNumberOfRace
    Else
        GetFirstHeatNumber = nTotalNum Mod N_NUMBER_OF_RACE
    End If

End Function

'
' ２組目人数算出
'
' １組目に回す人数によって２組目も変化する
'
' nTotalNum     IN      レースの総人数
' nMinNumberOfRace IN   組の最小人数
'
Function GetSecondHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer)

    If nTotalNum <= N_NUMBER_OF_RACE Then
        GetSecondHeatNumber = 0
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE = 0 Then
        GetSecondHeatNumber = N_NUMBER_OF_RACE
    ElseIf nTotalNum Mod N_NUMBER_OF_RACE <= nMinNumberOfRace Then
        GetSecondHeatNumber = N_NUMBER_OF_RACE + (nTotalNum Mod N_NUMBER_OF_RACE - nMinNumberOfRace)
    Else
        GetSecondHeatNumber = N_NUMBER_OF_RACE
    End If

End Function

'
' センターレーン算出
'
' nCount        IN      人数
' nStart        IN      開始位置
'
Function GetCenterLane(nCount As Integer, nStart As Integer, Optional bFlag As Boolean = True)
    If bFlag Then
        GetCenterLane = nStart + Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
    Else
        GetCenterLane = nStart + Application.WorksheetFunction.RoundDown((nCount) / 2, 0)
    End If
End Function

'
' 開始位置を算出
'
' nCount        IN      レース人数
' nCenterLane   IN      センターレーン
'
Function GetStartLane(nCount As Integer, nCenterLane As Integer, Optional bFlag As Boolean = True)
    If bFlag Then
        GetStartLane = nCenterLane - Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
    Else
        GetStartLane = nCenterLane - Application.WorksheetFunction.RoundDown((nCount) / 2, 0)
    End If
End Function

'
' レーン決定（単純方式）
'
' レーンは競技規則の単純方式で並べる
'
' nCenter       IN      センター
' nMax          IN      人数
' nOrder        IN      順番
' bFlag         In      True：通常／False：逆順
'
Function GetLane(nCenter As Integer, nMax As Integer, nOrder As Integer, Optional bFlag As Boolean = True)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    If bFlag Then
        GetLane = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        GetLane = nCenter + Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    End If
End Function

'
' レーン決定(平均分け方式)
'
' レーンは競技規則の平均分け方式で並べる
'
' nCenter       IN      センター
' nMax          IN      人数
' nOrder        IN      順番
'
Function GetLane2(nCenter As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = Application.WorksheetFunction.RoundUp((nMax - nOrder + 1) / N_AVERAGE_DEC_RACE, 0)
    GetLane2 = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
            * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
End Function

'
' 組番号決定(平均分け方式)
'
' プロNoの総人数と順番から組番号を算出する
'
' 組数を算出する
'
' nHeats        IN      組数
' nMax          IN      人数
' nOrder        IN      順番
'
Function GetOrderHeat(nHeats As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    GetOrderHeat = nHeats - (nNum - 1) Mod nHeats
End Function

'
' レーン決定(平均分け方式)
'
' レーンは競技規則の平均分け方式で並べる
'
' nStartRaceNo  IN/OUT  開始のRaceNo
' nStartHeat    IN      開始の組番号
' nMaxNum       IN      人数
' nOrder        IN      順番
' oProNo        IN      順番
' sTableName    IN      テーブル名
'
Sub AverageMethod(nStartRaceNo As Integer, nStartHeat As Integer, nMaxNum As Integer, _
nOrder As Integer, oProNo As Object, sTableName As String)
    
    Dim nCenterLane As Integer
    Dim nRow As Integer
    Dim nRaceNo As Integer
    Dim nHeat As Integer
    
    Dim nNum As Integer

    ' 組のセンター
    nCenterLane = GetCenterLane(N_NUMBER_OF_RACE, N_MIN_LANE_OF_RACE)
    
    ' 組の人数が残っている間
    For nNum = 1 To nMaxNum
        
        ' カレント行番号
        nRow = GetProNoRow(nOrder, oProNo)
        
        ' レースNo
        nRaceNo = (GetOrderHeat(N_AVERAGE_DEC_RACE, nMaxNum, nNum) + (nStartRaceNo / 10 - 1)) * 10
        ' 組番号
        nHeat = GetOrderHeat(N_AVERAGE_DEC_RACE, nMaxNum, nNum) + (nStartHeat - 1)
    
        ' レースNo、組の書込み
        Cells(nRow, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
        Cells(nRow, Range(sTableName & "[組]").Column).Value = nHeat
    
        ' レースNo、組、レーンを記述
        Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane2(nCenterLane, nMaxNum, nNum)
    
        ' 順番をインクリメント
        nOrder = nOrder + 1
    Next

    nStartRaceNo = (nStartRaceNo / 10 - 1 + N_AVERAGE_DEC_RACE) * 10

End Sub


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
    Call SheetActivate(S_ENTRY_SHEET_NAME)
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' 再ソート
    Call SortByRace(oWorkSheet, S_ENTRY_TABLE_NAME)
    
    ' レース番号修正
    Call SetRaceNo(oWorkSheet, S_ENTRY_TABLE_NAME)

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

