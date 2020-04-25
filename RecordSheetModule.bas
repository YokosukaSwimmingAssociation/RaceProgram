Attribute VB_Name = "RecordSheetModule"
'
' 種目名読込み
'
' 記録画面でProNoが入力されたら種目名を読込み表示する
' 存在しないProNoの場合は種目名は空欄となる
'
Sub 種目名読込み()
    Sheets("記録画面").Protect UserInterfaceOnly:=True
    For Each vNo In GetRange("プログラム種目番号")
        If vNo.Value = GetRange("記録画面種目番号").Value Then
            ' 種目区分と種目名を連結して表示する
            GetRange("記録画面種目名").Value = vNo.Offset(0, GetRange("Prog種目区分").Column - vNo.Column).Value _
                    & " " & vNo.Offset(0, GetRange("Prog種目名").Column - vNo.Column).Value
            Exit Sub
        End If
    Next vNo
    ' 該当がない場合は初期化
    Range("記録画面種目名").Value = ""
    Range("記録画面組").Value = 1
End Sub

'
' レース番号読込み
'
' 記録画面でProNoまたは組が入力されたらレース番号を読込み表示する
' 存在しない組み合わせの場合はレース番号は空欄となる
'
Sub レース番号読込み()
    Dim nProNo As Integer
    Dim nHeat As Integer

    nProNo = GetRange("記録画面種目番号").Value
    nHeat = GetRange("記録画面組").Value
    
    Dim sName As String
    sName = "プログラム組" & Format(nProNo, "0#") & "_" & Trim(Str(nHeat))

    If IsNameExists(sName) Then
        For Each vLane In Range(sName)
            If vLane.Offset(0, GetRange("HeaderレースNo").Column - vLane.Column).Value <> "" Then
                GetRange("記録画面レースNo").Value = vLane.Offset(0, GetRange("HeaderレースNo").Column - vLane.Column).Value
                Exit Sub
            End If
        Next vLane
    Else
        ' 存在しないプログラム番号の場合
        Range("記録画面レースNo").Value = ""
    End If
    
End Sub

'
' 選手名読込み
'
' 記録画面でレーンが入力された場合に選手名を読込み表示する
' レース番号が空欄の場合は何もしない
'
' oLaneCell     IN  変更のあったレーンのセル
'
Sub 選手名読込み(oLaneCell As Range)
    Dim nRaceNo As Integer
    nRaceNo = Range("記録画面レースNo").Value
    If nRaceNo = 0 Then
        Exit Sub
    End If
    
    Dim nLane As Integer
    nLane = oLaneCell.Value
    ' 選手名
    Cells(oLaneCell.Row, Range("記録画面選手名").Column).Value = SearchName(nRaceNo, nLane)
    ' チーム名
    Cells(oLaneCell.Row, Range("記録画面チーム名").Column).Value = SearchTeam(nRaceNo, nLane)
End Sub

'
' 選手名検索
'
' レース番号、レーン番号から選手名を検索する
'
' 名前「プログラムレースN」からレースのセルを取得して探索する
'
' nRaceNo           IN      レース番号
' nLane             IN      レーン番号
'
Function SearchName(nRaceNo As Integer, nLane As Integer)

    Dim sName As String
    sName = "プログラムレース" & Trim(Str(nRaceNo))

    ' 存在するレース番号の場合
    If IsNameExists(sName) Then
        ' レーン毎に処理する
        For Each vLaneNo In Range(sName)
            ' レーン番号が指定されたレーン番号の場合
            If vLaneNo.Offset(0, Range("Progレーン").Column - vLaneNo.Column).Value = nLane Then
                SearchName = vLaneNo.Offset(0, Range("Prog氏名").Column - vLaneNo.Column).Value
                ' 名前が空白用文字列の場合は空白にする
                If SearchName = S_BLANK_NAME Then
                    SearchName = ""
                End If
                Exit Function
            End If
        Next vLaneNo
    End If
    SearchName = ""
End Function

'
' チーム名検索
'
' レース番号、レーン番号からチーム名を検索する
'
' 名前「プログラムレースN」からレースのセルを取得して探索する
'
' nRaceNo           IN      レース番号
' nLane             IN      レーン番号
'
Function SearchTeam(nRaceNo As Integer, nLane As Integer)

    Dim sName As String
    sName = "プログラムレース" & Trim(Str(nRaceNo))

    ' 存在するレース番号の場合
    If IsNameExists(sName) Then
        ' レーン毎に処理する
        For Each vLaneNo In Range(sName)
            ' レーン番号が指定されたレーン番号の場合
            If vLaneNo.Offset(0, Range("Progレーン").Column - vLaneNo.Column).Value = nLane Then
                SearchTeam = vLaneNo.Offset(0, Range("Prog所属").Column - vLaneNo.Column).Value
                Exit Function
            End If
        Next vLaneNo
    End If
    SearchTeam = ""
End Function

'
' 違反反映
'
' oDqCell       IN  変更のあった違反セル
'
Sub 違反反映(oDqCell As Range)
    Application.EnableEvents = False
    
    Dim sDq As String
    sDq = STrimAll(oDqCell.Value)
    
    ' OPを設定されている場合
    If sDq = "OP" Then
        ' タイムを残してOPを設定
        Cells(oDqCell.Row, GetRange("記録画面備考").Column).Value = sDq
    
    ' 失格を設定されている場合
    ElseIf sDq <> "" Then
        ' タイムを空にして違反を設定
        Cells(oDqCell.Row, GetRange("記録画面タイム").Column).Value = ""
        Cells(oDqCell.Row, GetRange("記録画面備考").Column).Value = sDq
    
    ' 空白に戻した場合
    Else
        Call 大会記録判定(Cells(oDqCell.Row, GetRange("記録画面タイム").Column))
    End If

    Application.EnableEvents = True
End Sub

'
' 大会記録判定
'
' タイムが入力された場合に
'
' 名前「プログラムレースN」からレースのセルを取得して探索する
'
' oTimeCell       IN  変更のあったタイムセル
'
Sub 大会記録判定(oTimeCell As Range)
    Application.EnableEvents = False
    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim nTime As Long
    Dim nRecordTime As Long
    Dim nQualifyTime As Long

    nRaceNo = GetRange("記録画面レースNo").Value
   
    nLane = Cells(oTimeCell.Row, GetRange("記録画面レーン").Column).Value
    nTime = Cells(oTimeCell.Row, GetRange("記録画面タイム").Column).Value
    
    ' レーン、タイムに値が設定されている場合
    If nLane > 0 And nTime > 0 Then
        nRecordTime = SearchRecord(nRaceNo, nLane)
        nQualifyTime = SearchQualify(nRaceNo, nLane)
        If nQualifyTime > 0 And nTime > nQualifyTime Then
            ' 時間が標準記録より大きい場合はタイム失格
            Cells(oTimeCell.Row, GetRange("記録画面備考").Column).Value = "タイム失格"
        ElseIf nRecordTime = 0 Or nTime < nRecordTime Then
            ' 時間が大会記録より小さい場合は大会新（同一タイムはNG）
            Cells(oTimeCell.Row, GetRange("記録画面備考").Column).Value = "大会新"
        Else
            ' それ以外は空欄
            Cells(oTimeCell.Row, GetRange("記録画面備考").Column).Value = ""
        End If
    Else
        ' 何も入力されていないレーンも空欄
        Cells(oTimeCell.Row, GetRange("記録画面備考").Column).Value = ""
    End If
    Cells(oTimeCell.Row, GetRange("記録画面違反").Column).Value = ""

    Application.EnableEvents = True
End Sub

'
' 大会記録取得
'
' レース番号、レーン番号から大会記録を検索する
'
' 名前「プログラムレースN」からレースのセルを取得して探索する
'
' nRaceNo           IN      レース番号
' nLane             IN      レーン番号
'
Function SearchRecord(nRaceNo As Integer, nLane As Integer)

    Dim sName As String
    sName = "プログラムレース" & Trim(Str(nRaceNo))

    If IsNameExists(sName) Then
        For Each vLaneNo In Range(sName)
            If vLaneNo.Offset(0, GetRange("Progレーン").Column - vLaneNo.Column).Value = nLane Then
                If IsNumeric(vLaneNo.Offset(0, GetRange("Prog大会記録").Column - vLaneNo.Column).Value) Then
                    SearchRecord = CLng(vLaneNo.Offset(0, GetRange("Prog大会記録").Column - vLaneNo.Column).Value)
                Else
                    SearchRecord = 0
                End If
                Exit For
            End If
        Next vLaneNo
    End If
End Function

'
' 標準記録取得
'
' レース番号、レーン番号から標準記録を検索する
'
' 名前「プログラムレースN」からレースのセルを取得して探索する
'
' nRaceNo           IN      レース番号
' nLane             IN      レーン番号
'
Function SearchQualify(nRaceNo As Integer, nLane As Integer)

    Dim sName As String
    sName = "プログラムレース" & Trim(Str(nRaceNo))

    If IsNameExists(sName) Then
        For Each vLaneNo In Range(sName)
            If vLaneNo.Offset(0, GetRange("Progレーン").Column - vLaneNo.Column).Value = nLane Then
                If IsNumeric(vLaneNo.Offset(0, GetRange("Prog標準記録").Column - vLaneNo.Column).Value) Then
                    SearchQualify = CLng(vLaneNo.Offset(0, GetRange("Prog標準記録").Column - vLaneNo.Column).Value)
                Else
                    SearchQualify = 0
                End If
                Exit For
            End If
        Next vLaneNo
    End If
End Function

'
' 記録画面の初期化を行う
'
' レーンの初期化は行うが種目番号、組の初期化は行わない
'
Sub 初期化()
    Sheets("記録画面").Protect UserInterfaceOnly:=True
    
    ' イベント発生を抑制
    Call EventChange(False)
    
    For Each vLane In GetRange("記録画面レーン")
        vLane.Value = ""
        vLane.Offset(0, GetRange("記録画面タイム").Column - vLane.Column).Value = ""
        vLane.Offset(0, GetRange("記録画面選手名").Column - vLane.Column).Value = ""
        vLane.Offset(0, GetRange("記録画面チーム名").Column - vLane.Column).Value = ""
        vLane.Offset(0, GetRange("記録画面備考").Column - vLane.Column).Value = ""
        vLane.Offset(0, GetRange("記録画面違反").Column - vLane.Column).Value = ""
    Next vLane

    ' イベント発生を再開
    Call EventChange(True)
End Sub

'
' 入力データをプログラムに登録する
'
' 記録画面で登録ボタンが押された際にプログラムに記入する
'
Sub 登録()
    ' イベント発生を抑制
    Call EventChange(False)

    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim nTime As Long
    Dim sAdditional As String

    nRaceNo = GetRange("記録画面レースNo").Value
    
    For Each vLane In GetRange("記録画面レーン")
        nLane = Cells(vLane.Row, GetRange("記録画面レーン").Column).Value
        nTime = Cells(vLane.Row, GetRange("記録画面タイム").Column).Value
        sAdditional = Cells(vLane.Row, GetRange("記録画面備考").Column).Value
        
        If nLane <> 0 Then
            Call SetRecord(nRaceNo, nLane, nTime, sAdditional)
        End If
    Next vLane

    ' イベント発生を再開
    Call EventChange(True)
End Sub

'
' 入力データをプログラムに登録する
'
' nRaceNo           IN      レース番号
' nLane             IN      レーン番号
' nTime             IN      タイム
' sAdditional       IN      大会新
'
Function SetRecord(nRaceNo As Integer, nLane As Integer, nTime As Long, sAdditional As String)

    Dim sName As String
    sName = "プログラムレース" & Trim(Str(nRaceNo))

    For Each vLaneNo In GetRange(sName)
        If vLaneNo.Offset(0, GetRange("Progレーン").Column - vLaneNo.Column).Value = nLane Then
            If nTime = 0 Then
                ' タイムが入力されていない場合
                If sAdditional <> "" Then
                    ' 備考の値を設定
                    vLaneNo.Offset(0, GetRange("Prog備考").Column - vLaneNo.Column).Value = sAdditional
                Else
                    ' 備考が空欄なら棄権
                    vLaneNo.Offset(0, GetRange("Prog備考").Column - vLaneNo.Column).Value = "棄権"
                End If
            Else
                ' タイムが入力されている場合は時間と備考を設定
                vLaneNo.Offset(0, GetRange("Prog時間").Column - vLaneNo.Column).Value = nTime
                vLaneNo.Offset(0, GetRange("Prog備考").Column - vLaneNo.Column).Value = sAdditional
            End If
            Exit Function
        End If
    Next vLaneNo
End Function


'
' 順位決定
'
' 同一レースを考慮して、レースNoの中に含まれるプロNoに対して
' すべて順位をつける
'
Sub 順位決定()
    ' イベント発生を抑制
    Call EventChange(False)

    Dim nRaceNo As Integer
    nRaceNo = GetRange("記録画面レースNo").Value

    Dim sName As String
    sName = "プログラムレース" & Trim(Str(nRaceNo))

    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    Dim nProNo As Integer
    If IsNameExists(sName) Then
        For Each vRaceNo In GetRange(sName)
            nProNo = vRaceNo.Offset(0, GetRange("HeaderプロNo").Column - vRaceNo.Column).Value
            ' 最初の１回だけ実行
            If Not oProNo.Exists(nProNo) Then
                Call SetOrder(nProNo)
                oProNo.Add nProNo, 1
            End If
        Next vRaceNo
    End If

    ' イベント発生を再開
    Call EventChange(True)
End Sub

'
' 順番を決める
'
' nProNo            IN      種目番号
'
Sub SetOrder(nProNo As Integer)

    Dim sName As String
    sName = "プログラム番号" & Trim(Str(nProNo))
    
    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    ' 読込み
    Call ReadOrder(nProNo, sName, oProNo)

    ' ソートして順番を設定
    Call SortDictOrder(nProNo, sName, oProNo)

End Sub

'
' 順番を付けるレーン読み込む
'
' 記録がない場合は読み込まない
' タイム失格の場合は読み込まない
' OPの場合は読み込まない
'
' nProNo            IN      種目番号
' sName             IN      種目番号の名前
' oProNo            OUT     配列
'
Sub ReadOrder(nProNo As Integer, sName As String, oProNo As Object)
    
    ' 読込み
    Dim oSubClass As Object
    If IsNameExists(sName) Then
        ' レースNo毎に実施
        For Each vLane In Range(sName)
            ' 記録があり、タイム失格、OPでない合が対象
            If IsNumeric(vLane.Offset(0, GetRange("Prog時間").Column - vLane.Column).Value) And _
                vLane.Offset(0, GetRange("Prog備考").Column - vLane.Column).Value <> "タイム失格" And _
                vLane.Offset(0, GetRange("Prog備考").Column - vLane.Column).Value <> "OP" Then
                ' ソート区分（年齢区分）毎に順位をつける
                sSubClass = vLane.Offset(0, GetRange("Headerソート区分").Column - vLane.Column).Value
                If sSubClass = "" Then
                    ' ソート区分がない場合は１区分（ALL）としておく
                    sSubClass = "ALL"
                End If
                If Not oProNo.Exists(sSubClass) Then
                    Set oSubClass = CreateObject("Scripting.Dictionary")
                    oProNo.Add sSubClass, oSubClass
                End If

                ' Key（行）：Value（時間）として辞書型に登録
                oSubClass.Add vLane.Row, vLane.Offset(0, GetRange("Prog時間").Column - vLane.Column).Value
            End If
        Next vLane
    Else
        MsgBox "プログラム番号が不正です。"
        End
    End If
End Sub

'
' ソートして順番を設定
'
' nProNo            IN      種目番号
' sName             IN      種目番号の名前
' oProNo            OUT     配列
'
Sub SortDictOrder(nProNo As Integer, sName As String, oProNo As Object)
    ' 並び替え
    Dim oSubClass As Object
    Dim nOrder As Integer
    Dim nCount As Integer
    Dim nTime As Long
    Dim nPreTime As Long
    For Each vProNo In oProNo
        Set oSubClass = oProNo.Item(vProNo)
        ' 並び替えを実施
        Call DictQuickSort(oSubClass, "Value")
        nOrder = 1
        nCount = 1
        nPreTime = 0
        For Each vRow In oSubClass
            nTime = oSubClass.Item(vRow)
            ' 同一タイムでないときは順位を上げる
            If nTime > nPreTime Then
                nOrder = nCount
                nPreTime = nTime
            End If
            ' 順位を書き込む
            Sheets(Range(sName).Parent.Name).Cells(vRow, Range("Prog順位").Column).Value = nOrder
            nCount = nCount + 1
        Next vRow
    Next vProNo
End Sub

'
' 予選の場合に決勝を作成
'
Sub 決勝登録()
    ' イベント発生を抑制
    Call EventChange(False)

    ' 選手権以外は無効
    If GetRange("大会名").Value <> "横須賀選手権水泳大会" Then
        Exit Sub
    End If

    Dim nProNo As Integer
    nProNo = GetRange("記録画面種目番号").Value

    Dim nFinalNo As Integer
    nFinalNo = VLookupArea(nProNo, "選手権種目区分", "決勝番号")

    ' 決勝がない場合も無効
    If nProNo = nFileNo Then
        Exit Sub
    End If

    ' 決勝進出者を読み込む
    Dim oFinalist As Object
    Call ReadFinalist(nProNo, oFinalist)

    ' 大会記録と標準記録
    Dim nRecord As Long
    nRecord = VLookupArea(nFinalNo, "選手権大会記録", "記録")
    Dim nQualify As Long
    nQualify = VLookupArea(nProNo, "選手権種目区分", "標準記録")

    ' 決勝進出者を出力
    Call WriteFinalist(nFinalNo, oFinalist, nRecord, nQualify)

    ' イベント発生を再開
    Call EventChange(True)
End Sub

'
' 決勝進出者を読み込む
'
' nProNo            IN      種目番号
' oFinalist         OUT     決勝進出者の行番号配列
'
Sub ReadFinalist(nProNo As Integer, oFinalist As Object)
    Dim sName As String
    sName = "プログラム番号" & Trim(Str(nProNo))

    Set oFinalist = CreateObject("Scripting.Dictionary")

    If IsNameExists(sName) Then
        For Each vProNo In GetRange(sName)
            nOrder = GetOffset(vProNo, GetRange("Header順位").Column).Value
            ' 決勝人数まで保存
            If nOrder <= N_NUMBER_OF_RACE Then
                oFinalist.Add nOrder, vProNo
            End If
        Next vProNo
    End If
End Sub

'
' 決勝進出者を出力する
'
' nProNo            IN      種目番号
' oFinalist         IN      決勝進出者の行番号配列
' nRecord           IN      大会記録
' nQualify          IN      標準記録
'
Sub WriteFinalist(nProNo As Integer, oFinalist As Object, nRecord As Long, nQualify As Long)
    Dim sName As String
    sName = "プログラム番号" & Trim(Str(nProNo))

    Dim nLane As Integer
    Dim nOrder As Integer
    Dim nRow As Integer

    If IsNameExists(sName) Then
        For Each vProNo In GetRange(sName)
            ' レーン毎
            nLane = GetOffset(vProNo, GetRange("Headerレーン").Column).Value
            ' レーンから順位を取得
            nOrder = GetOrderByLane(GetCenterLane(N_NUMBER_OF_RACE, N_MIN_LANE_OF_RACE), nLane)
            ' 予選の行を取得
            Set vCell = oFinalist.Item(nOrder)
            
            GetOffset(vProNo, Range("Prog氏名").Column).Value = GetOffset(vCell, Range("Prog氏名").Column).Value
            GetOffset(vProNo, Range("Prog所属").Column).Value = GetOffset(vCell, Range("Prog所属").Column).Value
            GetOffset(vProNo, Range("Prog区分").Column).Value = GetOffset(vCell, Range("Prog区分").Column).Value
            GetOffset(vProNo, Range("Prog申込み記録").Column).Value = GetOffset(vCell, Range("Prog時間").Column).Value
            GetOffset(vProNo, Range("Prog大会記録").Column).Value = nRecord
            GetOffset(vProNo, Range("Prog標準記録").Column).Value = nQualify
            
        Next vProNo
    End If
End Sub

'
' レーンから順位を算出
'
' nCenterLane       IN      センターレーン
' nLane             IN      レーン番号
'
Function GetOrderByLane(nCenterLane As Integer, nLane As Integer)
    Dim nNum As Integer
    nNum = nLane - nCenterLane
    If nNum <= 0 Then
        GetOrderByLane = 2 * (1 - nNum) - 1
    Else
        GetOrderByLane = 2 * nNum
    End If
End Function
