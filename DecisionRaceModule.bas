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
    
    ' 組み合わせが学マだけ逆な対応（検証用の暫定）
    Dim bFlag As Boolean
    If GetRange("大会名").Value = "学童マスターズ大会" Then
        bFlag = False
    Else
        bFlag = True
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
    
    Dim nOrder As Integer           ' ProNo毎の順序（逆順）
    Dim nStartLane As Integer       ' 組の開始位置
    Dim nCenterLane As Integer      ' センターレーン
    Dim nMax As Integer             ' 組の中でLane決定する人数
    Dim nNum As Integer             ' Lene決定する人数の中の順位
    
    nRaceNo = 0
    
    ' プログラムNo毎
    For Each nProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(nProNo)
        
        ' プログラムNo毎の人数
        nNumOfProNo = GetNumberOfProNo(oProNo)
        
        ' 組数を算出
        nHeats = GetHeats(nNumOfProNo)
        Call GetNumberOfHeat(nNumOfProNo, nHeats, nNumOfHeats)
        
        ' ProNo毎の選手位置
        nOrder = 1
        
        ' 組毎
        For nHeat = 1 To nHeats
            ' レースNoをインクリメント
            nRaceNo = nRaceNo + 1
            
            ' 組の人数
            nNumOfHeat = nNumOfHeats(nHeat - 1)
            ' 組の残り人数
            nRemNumber = nNumOfHeat
        
            ' 組の開始位置
            nStartLane = GetStartLane(nNumOfHeat)
            
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
                    nCenterLane = GetCenterLane(nMax, nStartLane)
                    Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane3(nCenterLane, nMax, nNum, bFlag)
                
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
' nRow          IN      行数
' oProNo        IN      プロNo毎のエントリー配列
' nRaceNo       IN      レースNo
' nHeat         IN      組番号
' nNum          IN      順序
' nStartLane    IN      開始位置
'
Sub SetHeatLaneOrderRow(nRow As Integer, oProNo As Object, _
nRaceNo As Integer, nHeat As Integer, _
nNum As Integer, nStartLane As Integer)
    
    Dim sType As String             ' 種目区分
    
    ' レースNo、組の書込み
    Cells(nRow, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
    Cells(nRow, Range(sTableName & "[組]").Column).Value = nHeat
    
    ' 横須賀選手権水泳大会
    If GetRange("大会名").Value = "横須賀選手権水泳大会" Then
        Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
    ' 横須賀市民体育大会
    ElseIf GetRange("大会名").Value = "横須賀市民体育大会" Then
        ' 種目区分
        sType = Cells(nRow, Range(sTableName & "[種目区分]").Column).Value
              
        If sType = "年齢区分" Then
            ' 年齢区分
            Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane2(nHeat, oProNo.Count, nNum)
        Else
            ' 中学、高校
            Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane(nHeat, oProNo.Count, nNum)
        End If
    Else
        ' 横須賀マスターズ
        If Cells(nRow, Range(sTableName & "[ソート区分]").Column).Value <> "" Then
            Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane2(nHeat, oProNo.Count, nNum, False)
        ' 学童
        Else
            Cells(nRow, Range(sTableName & "[レーン]").Column).Value = GetLane(nHeat, oProNo.Count, nNum, False)
        End If
    End If
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
' nNumberOfHeat() OUT     組毎の人数配列
'
Sub GetNumberOfHeat(nTotalNum As Integer, nHeats As Integer, nNumberOfHeat() As Integer)
    
    ReDim nNumberOfHeat(nHeats - 1) As Integer
    
    ' １組目人数算出
    nNumberOfHeat(0) = GetFirstHeatNumber(nTotalNum)
    
    ' ２組目人数算出
    If nHeats >= 2 Then
        nNumberOfHeat(1) = GetSecondHeatNumber(nTotalNum)
    
        ' ３組目以降
        If nHeats > 2 Then
            For i = 2 To nHeats - 1
                nNumberOfHeat(i) = nNumberOfRace
            Next
        End If
    End If
End Sub

'
' 組番号決定
'
' プロNoの総人数と順番から組番号を算出する
'
' 組数、１組目の人数、２組目の人数を算出することで
' 組番号が算出可能
'
' nOrder        IN      順番
' nHeats        IN      組数
' nHeatNumber() IN      組毎の人数配列
'
Function GetOrderHeat(nOrder As Integer, nHeats As Integer, nHeatNumber() As Integer)

    ' １組目の場合
    If nOrder <= nHeatNumber(0) Then
        GetOrderHeat = 1
    ' ２組目の場合
    ElseIf nOrder <= nHeatNumber(0) + nHeatNumber(1) Then
        GetOrderHeat = 2
    ' ３組目以降の場合
    Else
        GetOrderHeat = GetHeats(nOrder - (nHeatNumber(0) + nHeatNumber(1))) + 2
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
' センターレーン算出
'
' nCount        IN      人数
' nStart        IN      開始位置
'
Function GetCenterLane(nCount As Integer, nStart As Integer)
     GetCenterLane = nStart + Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
End Function

'
' 開始位置を算出
'
' nCount        IN      レース人数
'
Function GetStartLane(nCount As Integer)
     GetStartLane = nCenterLane - Application.WorksheetFunction.RoundDown((nCount - 1) / 2, 0)
End Function

'
' レーン決定
'
' レーンは競技規則の単純方式で並べる
'
' nCenter       IN      センター
' nMax          IN      人数
' nOrder        IN      順番
' bFlag         In      True：通常／False：逆順
'
Function GetLane3(nCenter As Integer, nMax As Integer, nOrder As Integer, Optional bFlag As Boolean = True)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    If bFlag Then
        GetLane3 = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        GetLane3 = nCenter + Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
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

