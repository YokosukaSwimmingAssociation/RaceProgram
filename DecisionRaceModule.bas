Attribute VB_Name = "DecisionRaceModule"
Option Explicit    ''←変数の宣言を強制する

'
' 組み合わせ決定
'
' レースNoと組、レーンを設定する
'
Public Sub 組み合わせ決定()
    ' イベント発生を抑制
    Call EventChange(False)

    ' エクセルシートを選択
    Call SheetActivate(エントリーシート)

    ' 出力用ワークブック
    Dim oWorkBook As Workbook
    Set oWorkBook = ActiveWorkbook
    
    ' 出力用ワークシート
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet
    
    ' ProNo、ソート区分、申込み時間でソート
    Call SortByProNo(oWorkSheet, エントリーテーブル)
    
    ' 組み合わせ作成
    Call SetHeatLaneOrder(oWorkSheet, エントリーテーブル)
    
    ' レースNo, レーンでソート
    Call SortByRace(oWorkSheet, エントリーテーブル)
    
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
Private Sub SetHeatLaneOrder(oWorkSheet As Worksheet, sTableName As String)

    oWorkSheet.Activate
    
    ' エントリー一覧
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' データを格納
    Call ReadProNo(sTableName, oEntryList)
    
    ' 組、レーンを出力
    Call WriteHeatLaneOrder(sTableName, oEntryList)

End Sub

'
' プロNoをキーにデータ格納
'
' プロNo毎に通番を振ってその行番号を格納する
'
' sTableName    IN      テーブル名
' oEntryList    OUT     エントリー番号
'　└プロNo：
'　　└通番：ProNo行のセルオブジェクト
'
Private Sub ReadProNo(sTableName As String, oEntryList As Object)
    ' データを格納
    Dim vProNo As Variant
    Dim oProNo As Object
    For Each vProNo In Range(sTableName & "[プロNo]")
        If Not oEntryList.Exists(vProNo.Value) Then
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add vProNo.Value, oProNo
        End If
        oProNo.Add oProNo.Count + 1, vProNo
    Next vProNo
End Sub

'
' 組、レーンを出力する
'
' sTableName    IN      テーブル名
' oEntryList    IN      エントリー番号
'
Private Sub WriteHeatLaneOrder(sTableName As String, oEntryList As Object)

    Dim nNumOfProNo As Integer      ' プログラムNo毎の人数
    Dim nHeats As Integer           ' 組数
    Dim nNumOfHeats() As Integer    ' 組毎の人数
    Dim nMinNumberOfRace As Integer ' 組の最小人数
    nMinNumberOfRace = GetRange("大会組最少人数").Value
    Dim bAverage As Boolean         ' 平均分け方式の利用有無
    
    Dim nRaceNo As Integer          ' レースNo
    nRaceNo = 0
    
    ' プログラムNo毎
    Dim vProNo As Variant
    Dim oProNo As Object            ' プログラム毎の配列
    For Each vProNo In oEntryList.Keys
        Set oProNo = oEntryList.Item(vProNo)
        
        ' プログラムNo毎の人数
        nNumOfProNo = oProNo.Count
        
        ' 組数を算出
        nHeats = GetHeats(nNumOfProNo)
        Call GetNumberOfHeat(nNumOfProNo, nHeats, nNumOfHeats, nMinNumberOfRace)
        
        ' 平均分け方式を利用するか判定
        bAverage = IsAverageOrder(vProNo, nHeats)
        
        ' 組毎に計算
        Call WriteHeatLaneOrderByProNo(nRaceNo, nHeats, nNumOfHeats, oProNo, sTableName, bAverage)
    
    Next vProNo
End Sub

'
' プログラム番号毎に組、レーンを出力する
'
' nRaceNo       IN      レースNo
' nHeats        IN      組数
' nNumOfHeats   IN      組毎の人数
' oProNo        IN      プログラム毎の配列
' sTableName    IN      テーブル名
' bAverage      IN      平均分け方式の利用有無
'
Private Sub WriteHeatLaneOrderByProNo(nRaceNo As Integer, nHeats As Integer, nNumOfHeats() As Integer, _
oProNo As Object, sTableName As String, bAverage As Boolean)
    
    Dim nNumOfHeat As Integer       ' 組の人数
    Dim nOrder As Integer           ' ProNo毎の順序（逆順）
    nOrder = 1
    
    ' 組毎
    Dim nHeat As Integer
    For nHeat = 1 To nHeats
        ' レースNoをインクリメント(調整しやすいように10ずつ増やす)
        nRaceNo = nRaceNo + 10
        
        ' 平均分け方式
        If bAverage And nHeat = (nHeats - 平均分け組数) + 1 Then
            Call WriteHeatLaneOrderByAverage(nRaceNo, nHeat, nNumOfHeats, nOrder, oProNo, sTableName)
            Exit Sub
        Else
            ' 組の人数
            nNumOfHeat = nNumOfHeats(nHeat - 1)
            ' 組毎に出力
            Call WriteHeatLaneOrderByHeat(nRaceNo, nHeat, nNumOfHeat, nOrder, oProNo, sTableName)
        End If
    
    Next nHeat
End Sub

'
' 組毎に組、レーンを出力する
'
' nRaceNo       IN      レースNo
' nHeat         IN      組番号
' nNumOfHeat    IN      組の人数
' nOrder        IN      順番
' oProNo        IN      順番
' sTableName    IN      テーブル名
'
Private Sub WriteHeatLaneOrderByHeat(nRaceNo As Integer, nHeat As Integer, nNumOfHeat As Integer, _
nOrder As Integer, oProNo As Object, sTableName As String)
    
    ' 組み合わせが学マだけ逆な対応（検証用の暫定）
    Dim bFlag As Boolean
    If GetRange("大会名").Value = 学マ大会 Then
        bFlag = False
    Else
        bFlag = True
    End If
    
    Dim nStartLane As Integer       ' 組の開始位置
    Dim nTargetNum As Integer             ' 組の中でLane決定する人数
    
    Dim nNumOfSortClass As Integer  ' ソート区分毎の残り人数
    Dim nRemNumber As Integer       ' 組の残り人数
    nRemNumber = nNumOfHeat
    
    ' 組の開始位置
    nStartLane = GetStartLane(nNumOfHeat, GetCenterLane(Range("大会組レース定員").Value, GetRange("大会組最小レーン番号").Value), bFlag)
    
    ' 組の人数が残っている間
    While nRemNumber > 0
        ' ソート区分毎の残り人数
        nNumOfSortClass = GetNumberOfSortClass(nOrder, oProNo, Range(sTableName & "[ソート区分]").Column)
        If nNumOfSortClass <= nRemNumber Then
            ' Lane決定する人数
            nTargetNum = nNumOfSortClass
        Else
            nTargetNum = nRemNumber
        End If
        
        ' ソート区分毎に組、レーンを出力
        Call WriteHeatLaneOrderBySortClass(nRaceNo, nHeat, nTargetNum, nOrder, oProNo, sTableName, nStartLane, bFlag)
    
        ' 開始位置を変更
        nStartLane = nStartLane + nTargetNum
    
        ' 残り人数を減算
        nRemNumber = nRemNumber - nTargetNum
    Wend
End Sub

'
' ソート区分毎に組、レーンを出力する
'
' nRaceNo       IN      レースNo
' nHeat         IN      組番号
' nTargetNum    IN      対象の人数
' nOrder        IN/OUT  順番
' oProNo        IN      順番の配列
' sTableName    IN      テーブル名
' nStartLane    IN      開始レーン
' bFlag         In      True：通常／False：逆順
'
Private Sub WriteHeatLaneOrderBySortClass(nRaceNo As Integer, nHeat As Integer, _
nTargetNum As Integer, ByRef nOrder As Integer, _
oProNo As Object, sTableName As String, _
nStartLane As Integer, Optional bFlag As Boolean = True)

    Dim nCenterLane As Integer      ' センターレーン

    ' レーンを決定する
    Dim oCell As Range              ' カレント行のセル
    Dim nIndex As Integer           ' Lene決定する人数の中の順位
    For nIndex = 1 To nTargetNum
        ' カレント行番号
        Set oCell = oProNo.Item(nOrder)
    
        ' レースNo、組の書込み
        GetOffset(oCell, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
        GetOffset(oCell, Range(sTableName & "[組]").Column).Value = nHeat
    
        ' レースNo、組、レーンを記述
        nCenterLane = GetCenterLane(nTargetNum, nStartLane, bFlag)
        GetOffset(oCell, Range(sTableName & "[レーン]").Column).Value = GetLane(nCenterLane, nTargetNum, nIndex, bFlag)
    
        ' 順番をインクリメント
        nOrder = nOrder + 1
    Next nIndex

End Sub

'
' 開始位置から同一ソート区分の件数
'
' 開始位置で指定されたソート区分と同じ値の間はカウントする
'
' nIndex            IN      開始位置
' oProNo            IN      プロNo毎のエントリー配列
' nSortClassColumn  IN      ソート区分のカラム位置
'
Private Function GetNumberOfSortClass(nIndex As Integer, oProNo As Object, nSortClassColumn As Integer) As Integer
    
    GetNumberOfSortClass = 1
    
    Dim vProNo As Object
    Set vProNo = oProNo.Item(nIndex)
    Dim sSortClass As String
    sSortClass = GetOffset(vProNo, nSortClassColumn).Value
    
    Dim i As Integer
    For i = nIndex + 1 To oProNo.Count
        Set vProNo = oProNo.Item(i)
        If GetOffset(vProNo, nSortClassColumn).Value = sSortClass Then
            GetNumberOfSortClass = GetNumberOfSortClass + 1
        Else
            Exit Function
        End If
    Next i
End Function

'
' 組人数配列を設定
'
' nTotalNum     IN      プロNoのエントリー数
' nHeat         IN      組数
' nNumberOfHeat() OUT   組毎の人数配列
' nMinNumberOfRace IN   組の最小人数
'
Private Sub GetNumberOfHeat(nTotalNum As Integer, nHeats As Integer, nNumberOfHeat() As Integer, nMinNumberOfRace As Integer)
    
    ReDim nNumberOfHeat(nHeats - 1) As Integer
    
    ' １組目人数算出
    nNumberOfHeat(0) = GetFirstHeatNumber(nTotalNum, nMinNumberOfRace)
    
    ' ２組目人数算出
    If nHeats >= 2 Then
        nNumberOfHeat(1) = GetSecondHeatNumber(nTotalNum, nMinNumberOfRace)
    
        ' ３組目以降
        If nHeats > 2 Then
            Dim i As Integer
            For i = 2 To nHeats - 1
                nNumberOfHeat(i) = Range("大会組レース定員").Value
            Next i
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
Private Function GetHeats(nTotalNum As Integer) As Integer

    GetHeats = Application.WorksheetFunction.RoundUp(nTotalNum / Range("大会組レース定員").Value, 0)

End Function

'
' １組目人数算出
'
' 総人数が最小人数以上いる場合、最小人数は１組目に残す
'
' nTotalNum     IN      レースの総人数(複数組の総数)
' nMinNumberOfRace IN   組の最小人数
'
Private Function GetFirstHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer) As Integer

    Dim maxNum
    maxNum = Range("大会組レース定員").Value
    If nTotalNum <= maxNum Then
        GetFirstHeatNumber = nTotalNum
    ElseIf nTotalNum Mod maxNum = 0 Then
        GetFirstHeatNumber = maxNum
    ElseIf nTotalNum Mod maxNum <= nMinNumberOfRace Then
        GetFirstHeatNumber = nMinNumberOfRace
    Else
        GetFirstHeatNumber = nTotalNum Mod maxNum
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
Private Function GetSecondHeatNumber(nTotalNum As Integer, nMinNumberOfRace As Integer) As Integer

    Dim maxNum
    maxNum = Range("大会組レース定員").Value
    If nTotalNum <= maxNum Then
        GetSecondHeatNumber = 0
    ElseIf nTotalNum Mod maxNum = 0 Then
        GetSecondHeatNumber = maxNum
    ElseIf nTotalNum Mod maxNum <= nMinNumberOfRace Then
        GetSecondHeatNumber = maxNum + (nTotalNum Mod maxNum - nMinNumberOfRace)
    Else
        GetSecondHeatNumber = maxNum
    End If

End Function

'
' センターレーン算出
'
' nCount        IN      人数
' nStart        IN      開始位置
'
Public Function GetCenterLane(nCount As Integer, nStart As Integer, Optional bFlag As Boolean = True) As Integer
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
Private Function GetStartLane(nCount As Integer, nCenterLane As Integer, Optional bFlag As Boolean = True) As Integer
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
' 逆順にするために総人数から順序を引いている
'
' nCenter       IN      センター
' nMaxNum       IN      総人数
' nOrder        IN      順番
' bFlag         In      True：通常／False：逆順
'
Private Function GetLane(nCenter As Integer, nMaxNum As Integer, nOrder As Integer, Optional bFlag As Boolean = True)
    Dim nNum As Integer
    nNum = nMaxNum - nOrder + 1
    If bFlag Then
        GetLane = nCenter - Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    Else
        GetLane = nCenter + Application.WorksheetFunction.Power(-1, nNum - 1) _
                * Application.WorksheetFunction.RoundUp((nNum - 1) / 2, 0)
    End If
End Function

'
'  平均分け方式を採用するか
'
' 選手権大会の予選で、混合分け方式を指定されている場合のみ対象
'
' vProNo        IN      プログラム番号
' nHeats        IN      組数
'
Private Function IsAverageOrder(vProNo As Variant, nHeats As Integer) As Boolean
    IsAverageOrder = False
    If GetRange("大会名").Value = 選手権大会 And nHeats >= 平均分け組数 Then
        If GetRange("大会組合せ方式").Value = "混合分け方式" And _
            VLookupArea(vProNo, "選手権種目区分", "予選／決勝") = "予選" Then
            IsAverageOrder = True
        End If
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
Private Function GetLane2(nCenter As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = Application.WorksheetFunction.RoundUp((nMax - nOrder + 1) / 平均分け組数, 0)
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
Private Function GetOrderHeat(nHeats As Integer, nMax As Integer, nOrder As Integer)
    Dim nNum As Integer
    nNum = nMax - nOrder + 1
    GetOrderHeat = nHeats - (nNum - 1) Mod nHeats
End Function


'
' レーン決定(平均分け方式)
'
' nStartRaceNo  IN/OUT  開始のRaceNo
' nStartHeat    IN      開始の組番号
' nNumOfHeats   IN      組毎の人数
' nOrder        IN      順番
' oProNo        IN      順番
' sTableName    IN      テーブル名
'
Private Sub WriteHeatLaneOrderByAverage(nStartRaceNo As Integer, nStartHeat As Integer, _
nNumOfHeats() As Integer, nOrder As Integer, oProNo As Object, sTableName As String)
    Dim nMaxNum As Integer
    nMaxNum = 0
    Dim i As Integer
    For i = nStartHeat To nStartHeat + 平均分け組数 - 1
        nMaxNum = nMaxNum + nNumOfHeats(i - 1)
    Next i
    Call AverageMethod(nStartRaceNo, nStartHeat, nMaxNum, nOrder, oProNo, sTableName)
    nStartRaceNo = (nStartRaceNo / 10 - 1 + 平均分け組数) * 10
End Sub


'
' レーン決定(平均分け方式)
'
' レーンは競技規則の平均分け方式で並べる
'
' nStartRaceNo  IN      開始のRaceNo
' nStartHeat    IN      開始の組番号
' nMaxNum       IN      人数
' nOrder        IN      順番
' oProNo        IN      順番配列
' sTableName    IN      テーブル名
'
Private Sub AverageMethod(nStartRaceNo As Integer, nStartHeat As Integer, nMaxNum As Integer, _
nOrder As Integer, oProNo As Object, sTableName As String)
    
    Dim nCenterLane As Integer
    Dim nRaceNo As Integer
    Dim nHeat As Integer
    
    ' 組のセンター
    nCenterLane = GetCenterLane(Range("大会組レース定員").Value, GetRange("大会組最小レーン番号").Value)
    
    ' 組の人数が残っている間
    Dim oCell As Range              ' カレント行のセル
    Dim nIndex As Integer
    For nIndex = 1 To nMaxNum
        
        ' カレント行番号
        Set oCell = oProNo.Item(nOrder)
        
        ' レースNo
        nRaceNo = (GetOrderHeat(平均分け組数, nMaxNum, nIndex) + (nStartRaceNo / 10 - 1)) * 10
        ' 組番号
        nHeat = GetOrderHeat(平均分け組数, nMaxNum, nIndex) + (nStartHeat - 1)
    
        ' レースNo、組の書込み
        GetOffset(oCell, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
        GetOffset(oCell, Range(sTableName & "[組]").Column).Value = nHeat
    
        ' レースNo、組、レーンを記述
        GetOffset(oCell, Range(sTableName & "[レーン]").Column).Value = GetLane2(nCenterLane, nMaxNum, nIndex)
    
        ' 順番をインクリメント
        nOrder = nOrder + 1
    Next nIndex
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
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = SheetActivate(エントリーシート)
    
    ' 再ソート
    Call SortByRace(oWorkSheet, エントリーテーブル)
    
    ' ProNo、組の重複チェック
    Dim oEntryList As Object
    Call ReadEntrySheet(エントリーテーブル, oEntryList)
    
    ' レース番号修正
    If GetRange("大会名").Value = 選手権大会 Then
        Call ModifyRaceNoForSenshuken(エントリーテーブル)
    Else
        Call ModifyRaceNo(エントリーテーブル)
    End If

    ' イベント発生を再開
    Call EventChange(True)
End Sub

'
' レース番号修正
'
' sTableName    IN  テーブル名
'
Private Sub ModifyRaceNo(sTableName As String)
    
    ' エントリー一覧
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' データを格納＆レース番号のチェック
    Call ReadEntryByRaceNo(sTableName, oEntryList)
    
    ' 出力
    Call ModifyEntryByRaceNo(sTableName, oEntryList)

End Sub

'
' レース番号の読込み
'
' 読込みながらレース番号中のレーン重複チェックも行う
'
' sTableName    IN      テーブル名
' oEntryList    IN/OUT  エントリー一覧
'   └レースNo
'   　　└レーン：プロNo列のセル
'
Private Sub ReadEntryByRaceNo(sTableName As String, oEntryList As Object)
    Dim nLane As Integer
    Dim oRaceNo As Object
    Dim vRaceNo As Variant
    For Each vRaceNo In Range(sTableName & "[レースNo]")
        ' 存在しないレースNoの場合は
        If Not oEntryList.Exists(vRaceNo.Value) Then
            ' エントリー一覧に登録する
            Set oRaceNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add vRaceNo.Value, oRaceNo
        End If
        
        ' レーン番号を取得
        nLane = GetOffset(vRaceNo, Range(sTableName & "[レーン]").Column).Value
        
        ' レース番号に対するレーンの重複チェック
        If oRaceNo.Exists(nLane) Then
            MsgBox "レースNo：" & Str(vRaceNo.Value) & vbCrLf & _
                    "レーン　：" & Str(nLane) & vbCrLf & _
                    "が重複しています。"
            Range(sTableName).Parent.Activate
            Range(vRaceNo, GetOffset(vRaceNo, Range(sTableName & "[レーン]").Column)).Select
            vRaceNo.Activate
            End
        Else
            oRaceNo.Add nLane, vRaceNo
        End If
    Next vRaceNo
End Sub

'
' レース番号を修正出力
'
' sTableName    IN      テーブル名
' oEntryList    IN/OUT  エントリー一覧
'
Private Sub ModifyEntryByRaceNo(sTableName As String, oEntryList As Object)
    ' レース番号修正
    Dim nRaceNo As Integer
    nRaceNo = 1
    
    Dim oCell As Range
    Dim vRaceNo As Variant
    Dim oRaceNo As Object
    For Each vRaceNo In oEntryList.Keys
        Set oRaceNo = oEntryList.Item(vRaceNo)
        
        ' レーン
        Dim vLane As Variant
        For Each vLane In oRaceNo.Keys
            Set oCell = oRaceNo.Item(vLane)
            
            ' レース番号
            GetOffset(oCell, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
        Next
        nRaceNo = nRaceNo + 1
    Next vRaceNo
End Sub

'
' レース番号修正(選手権用)
'
' sTableName    IN  テーブル名
'
Private Sub ModifyRaceNoForSenshuken(sTableName As String)
    
    ' エントリー一覧
    Dim oEntryList As Object
    Set oEntryList = CreateObject("Scripting.Dictionary")
    
    ' データを格納＆レース番号のチェック
    Call ReadRaceForSenshuken(sTableName, oEntryList)
    
    ' 出力
    Call ModifyEntryByProNoForSenshuken(sTableName, oEntryList)

End Sub

'
' レース番号の読込み(選手権用)
'
' 読込みながらレース番号中のレーン重複チェックも行う
'
' sTableName    IN      テーブル名
' oEntryList    IN/OUT エントリー一覧
' └プロNo
'   └レースNo
'   　　└レーン：行
'
Private Sub ReadRaceForSenshuken(sTableName As String, oEntryList As Object)
    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim oProNo As Object
    Dim oRaceNo As Object
    
    Dim vProNo As Variant
    For Each vProNo In Range(sTableName & "[プロNo]")
        ' 存在しないプロNoの場合は
        If Not oEntryList.Exists(vProNo.Value) Then
            ' エントリー一覧に登録する
            Set oProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add vProNo.Value, oProNo
        End If
        
        ' 存在しないレースNoの場合は
        nRaceNo = GetOffset(vProNo, Range(sTableName & "[レースNo]").Column).Value
        If Not oProNo.Exists(nRaceNo) Then
            ' エントリー一覧に登録する
            Set oRaceNo = CreateObject("Scripting.Dictionary")
            oProNo.Add nRaceNo, oRaceNo
        End If
        
        ' レーン番号を取得
        nLane = GetOffset(vProNo, Range(sTableName & "[レーン]").Column).Value
        ' レース番号に対するレーンの重複チェック
        If oRaceNo.Exists(nLane) Then
            MsgBox "レースNo：" & Str(nRaceNo) & vbCrLf & _
                    "レーン　：" & Str(nLane) & vbCrLf & _
                    "が重複しています。"
            Range(sTableName).Parent.Activate
            Range(vProNo, GetOffset(vProNo, Range(sTableName & "[レーン]").Column)).Select
            vProNo.Activate
            End
        Else
            oRaceNo.Add nLane, vProNo
        End If
    Next vProNo
End Sub

'
' レース番号修正(選手権用)
'
' sTableName    IN  テーブル名
' oEntryList    IN  エントリー一覧
'
Private Sub ModifyEntryByProNoForSenshuken(sTableName As String, oEntryList As Object)
    
    ' レース番号修正
    Dim nRaceNo As Integer
    nRaceNo = 1
    
    Dim vProNo As Variant
    For Each vProNo In GetAreaKeyData("選手権種目区分")
        ' ProNoのエントリーがない場合はスキップ
        If oEntryList.Exists(vProNo.Value) Then
            
            ' 予選決勝の修正
            Call ModifyFinalEntry(vProNo, oEntryList, nRaceNo)
            
            ' レースNoの修正
            If oEntryList.Exists(vProNo.Value) Then
                Call ModifyEntryByRaceNoForSenshuken(sTableName, vProNo, oEntryList, nRaceNo)
            End If
                 
        End If
    Next vProNo

End Sub

'
' 予選決勝の修正
'
' vProNo        IN      ProNo列のセル
' oEntryList    IN/OUT  エントリー配列
' nRaceNo       IN/OUT  レース番号
'
Private Sub ModifyFinalEntry(vProNo As Variant, oEntryList As Object, nRaceNo As Integer)
    ' ProNoのエントリーを取得
    Dim oProNo As Object
    Set oProNo = oEntryList.Item(vProNo.Value)
    
    ' 決勝番号を取得
    Dim nFinalNo As Integer
    nFinalNo = VLookupArea(vProNo.Value, "選手権種目区分", "決勝番号")
    
    ' 予選の場合
    If vProNo.Value <> nFinalNo Then
        ' 組が１組の場合
        If oProNo.Count = 1 Then
            ' ProNoを決勝に入れ替える
            oEntryList.Add nFinalNo, oProNo
            oEntryList.Remove vProNo.Value
            ' レース番号は振らない
            Set oProNo = CreateObject("Scripting.Dictionary")
        ElseIf oProNo.Count > 1 Then
            ' 決勝に空のエントリーを作っておく
            Dim oFinalProNo As Object
            Set oFinalProNo = CreateObject("Scripting.Dictionary")
            oEntryList.Add nFinalNo, oFinalProNo
        End If
    Else
        ' 決勝の組が0の場合
        If oProNo.Count = 0 Then
            ' 予選決勝があるパターンなのでインクリメントしておく
            nRaceNo = nRaceNo + 1
        End If
    End If
End Sub

'
' レース番号の修正出力
'
' sTableName    IN  テーブル名
' vProNo        IN      ProNo列のセル
' oEntryList    IN/OUT  エントリー配列
' nRaceNo       IN/OUT  レース番号
'
Private Sub ModifyEntryByRaceNoForSenshuken(sTableName As String, _
vProNo As Variant, oEntryList As Object, nRaceNo As Integer)
    Dim oCell As Range
    Dim oProNo As Object
    Dim oRaceNo As Object
    
    ' ProNoのエントリーを取得
    Set oProNo = oEntryList.Item(vProNo.Value)
    
    Dim vRaceNo As Variant
    For Each vRaceNo In oProNo.Keys()
        Set oRaceNo = oProNo.Item(vRaceNo)
    
        ' レーン
        Dim vLane As Variant
        For Each vLane In oRaceNo.Keys
            Set oCell = oRaceNo.Item(vLane)
            
            ' レース番号
            GetOffset(oCell, Range(sTableName & "[レースNo]").Column).Value = nRaceNo
        Next
    
        nRaceNo = nRaceNo + 1
    Next vRaceNo
End Sub


'
' レースNo、組でソートする
'
' oWorkSheet    IN      ワークシート
' sTableName    IN      テーブル名
'
Private Sub SortByRace(oWorkSheet As Worksheet, sTableName As String)

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

