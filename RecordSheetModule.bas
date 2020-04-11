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
Sub 選手名読込み()
    Dim nRaceNo As Integer
    nRaceNo = Range("記録画面レースNo").Value
    If nRaceNo = 0 Then
        Exit Sub
    End If
    
    Dim nLane As Integer
    For Each vLane In Range("記録画面レーン")
        nLane = Cells(vLane.Row, Range("記録画面レーン").Column).Value
        ' 選手名
        Cells(vLane.Row, Range("記録画面選手名").Column).Value = SearchName(nRaceNo, nLane)
        ' チーム名
        Cells(vLane.Row, Range("記録画面チーム名").Column).Value = SearchTeam(nRaceNo, nLane)
    Next vLane
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
                If SearchName = sBlankName Then
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
' 大会記録判定
'
' タイムが入力された場合に
'
' 名前「プログラムレースN」からレースのセルを取得して探索する
'
Sub 大会記録判定()
    Dim nRaceNo As Integer
    Dim nLane As Integer
    Dim nTime As Long
    Dim nRecordTime As Integer

    nRaceNo = GetRange("記録画面レースNo").Value
   
    For Each vLane In GetRange("記録画面レーン")
        nLane = Cells(vLane.Row, GetRange("記録画面レーン").Column).Value
        nTime = Cells(vLane.Row, GetRange("記録画面タイム").Column).Value
        
        If nLane > 0 And nTime > 0 Then
            If nTime < SearchRecord(nRaceNo, nLane) Then
                ' 時間が大会記録より小さい場合は大会新（同一タイムはNG）
                Cells(vLane.Row, GetRange("記録画面大会新").Column).Value = "大会新"
            Else
                ' それ以外は空欄
                Cells(vLane.Row, GetRange("記録画面大会新").Column).Value = ""
            End If
        Else
            ' 何も入力されていないレーンも空欄
            Cells(vLane.Row, GetRange("記録画面大会新").Column).Value = ""
        End If
    Next vLane
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
                SearchRecord = vLaneNo.Offset(0, GetRange("Prog大会記録").Column - vLaneNo.Column).Value
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
        vLane.Offset(0, GetRange("記録画面大会新").Column - vLane.Column).Value = ""
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
        sAdditional = Cells(vLane.Row, GetRange("記録画面大会新").Column).Value
        
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
                ' タイムが入力されていない場合は棄権
                vLaneNo.Offset(0, GetRange("Prog備考").Column - vLaneNo.Column).Value = sAdditional
                vLaneNo.Offset(0, GetRange("Prog備考").Column - vLaneNo.Column).Value = "棄権"
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

    ' ソート
    Call SortDictOrder(nProNo, sName, oProNo)

End Sub

'
' 順番を読み込む
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
            ' 時間が入力されている場合が対象
            If IsNumeric(vLane.Offset(0, GetRange("Prog時間").Column - vLane.Column).Value) Then
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
' 順番を読み込む
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
