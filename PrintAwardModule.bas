Attribute VB_Name = "PrintAwardModule"
Option Explicit    ''←変数の宣言を強制する
'
' 賞状印刷
'
' 指定したレースNoに存在するProNoの賞状を印刷する
'
Public Sub 賞状印刷()

    ' イベント発生を抑制
    Call EventChange(False)

    ' 記録画面シートを保存
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    ' レース番号
    Dim nRaceNo As Integer
    nRaceNo = GetRange("記録画面レースNo").Value

    ' 賞状印刷
    Call PrintAwardByRace(nRaceNo)

    ' 記録画面シートに戻る
    oWorkSheet.Activate

    ' イベント発生を再開
    Call EventChange(True)

End Sub

'
' レース番号指定賞状印刷
'
' nRaceNo           IN      レースNo
'
Private Function PrintAwardByRace(nRaceNo As Integer)

    ' レースの名前取得
    Dim sName As String
    sName = "プログラムレース" & Trim(Str(nRaceNo))
    If Not IsNameExists(sName) Then
        MsgBox "印刷対象が取得できませんでした。" & vbCrLf & _
                "レースNoが正しく指定されているか確認してください。" & vbCrLf & _
                "正しく指定されている場合は、プログラム名前定義を実行してみてください。", vbOKOnly
        End
    End If

    ' 賞状印刷
    Call PrintAwardByName(sName)

End Function

'
' 名前指定賞状印刷
'
' sName             IN      レースの名前定義
'
Private Function PrintAwardByName(sName As String)

    Dim oProNo As Object
    Set oProNo = CreateObject("Scripting.Dictionary")

    Dim nProNo As Integer
    Dim vRaceNo As Variant
    For Each vRaceNo In GetRange(sName)
        nProNo = GetOffset(vRaceNo, GetRange("HeaderプロNo").Column).Value
        ' 最初の１回だけ実行
        If Not oProNo.Exists(nProNo) Then
            ' 賞状印刷対象か確認
            If CheckTarget(nProNo) Then
                ' 賞状を印刷
                Call PrintAwardByProNo(nProNo)
            End If
            oProNo.Add nProNo, 1
        End If
    Next vRaceNo

End Function


'
' 印刷対象かを確認する
'
' 学マ大会は学童のみ対象（種目区分から対象を取得）
' 市民大会はすべて対象
' 選手権大会は決勝のみ対象（種目区分から対象を取得）
'
'
' nProNo            IN      プロNo
'
Private Function CheckTarget(nProNo As Integer) As Boolean

    ' 大会名
    Dim sGameName As String
    sGameName = GetRange("大会名").Value

    If sGameName = 選手権大会 Then
        ' 予選・決勝の確認
        If VLookupArea(nProNo, "選手権種目区分", "予選／決勝") <> "予選" Then
            CheckTarget = True
        Else
            CheckTarget = False
        End If
    ElseIf sGameName = 市民大会 Then
        CheckTarget = True
    ElseIf sGameName = 学マ大会 Then
        ' 種目の大会を取得
        If VLookupArea(nProNo, "学マ種目区分", "大会区分") = "学童" Then
            CheckTarget = True
        Else
            CheckTarget = False
        End If
    Else
        MsgBox "大会名が正しく指定されていません。", vbOKOnly
        End
    End If

End Function

'
' 賞状を印刷する
'
' 指定したプロNoの中で1位～3位の賞状を印刷する
'
' nProNo            IN      プロNo
'
Private Sub PrintAwardByProNo(nProNo As Integer)

    ' 大会名
    Dim sGameName As String
    sGameName = GetRange("大会名").Value
    
    ' 種目区分を取得
    Dim sMasterName As String
    sMasterName = GetMaster(GetRange("大会名").Value)
    
    Dim sRaceClass As String ' 区分
    sRaceClass = VLookupArea(nProNo, sMasterName, "種目区分")
    Dim sGender As String ' 性別
    sGender = VLookupArea(nProNo, sMasterName, "性別")
    Dim sDistance As String ' 距離
    sDistance = Replace(VLookupArea(nProNo, sMasterName, "距離"), "M", "")
    Dim sStyle As String ' 種目
    sStyle = VLookupArea(nProNo, sMasterName, "種目")
    Dim nMaxOrder As Integer ' 出力する順位
    nMaxOrder = VLookupArea(sGameName, "設定各種", "賞状順位")
    
    Dim sName As String
    sName = "プログラム番号" & Trim(CStr(nProNo))
    
    Dim vProNo As Variant
    Dim nOrder As Integer
    For Each vProNo In GetRange(sName)
        nOrder = Val(GetOffset(vProNo, GetRange("Prog順位").Column).Value)
        If nOrder >= 1 And nOrder <= nMaxOrder Then
            Call PrintAwardByLine(sGameName, vProNo, sRaceClass, sGender, sDistance, sStyle)
        End If
    Next vProNo
    
End Sub

'
' 行指定で賞状を印刷する
'
' 指定した行のレコードを印刷する
'
' sGameName         IN      大会名
' vProNo            IN      ProNo
' sRaceClass        IN      種目区分
' sGender           IN      性別
' sDistance         IN      距離
' sStyle            IN      種目
'
Private Sub PrintAwardByLine(sGameName As String, _
vProNo As Variant, _
sRaceClass As String, _
sGender As String, _
sDistance As String, _
sStyle As String)
   
    ' 共通の設定
    GetRange("賞状順位").Value = GetOffset(vProNo, Range("Prog順位").Column).Value
    GetRange("賞状タイム").Value = GetOffset(vProNo, Range("Prog時間").Column).Value
    If GetOffset(vProNo, Range("Prog備考").Column).Value = "大会新" Then
        GetRange("賞状大会新").Value = "大会新"
    Else
        GetRange("賞状大会新").Value = ""
    End If
    GetRange("賞状氏名").Value = GetOffset(vProNo, Range("Prog氏名").Column).Value
    GetRange("賞状所属").Value = GetOffset(vProNo, Range("Prog所属").Column).Value
   
    If GetRange("賞状タイム").Value >= 10000 Then
        GetRange("賞状タイム").NumberFormatLocal = "#""分""##""秒""##"
    Else
        GetRange("賞状タイム").NumberFormatLocal = "##""秒""##"
    End If

    ' 大会固有の設定
    If sGameName = 選手権大会 Then
        Call SetAwardValForSenshuken(sGender, sDistance, sStyle)
    ElseIf sGameName = 市民大会 Then
        Call SetAwardValForShimin(sRaceClass, sGender, sDistance, sStyle, _
                GetOffset(vProNo, Range("Prog区分").Column).Value)
    Else
        Call SetAwardValForGakudo(sRaceClass, sGender, sDistance, sStyle)
    End If

    ' 印刷
    Call PrintAward

End Sub

'
' 賞状印刷
'
Private Sub PrintAward()
    ' プレビュー有無
    Dim bPreview As Boolean
    If GetRange("大会印刷プレビュー").Value = "する" Then
        bPreview = True
    Else
        bPreview = False
    End If
    
    ' プリンタ名
    Dim sPrinterName As String
    sPrinterName = GetRange("プリンタ名").Value
    If sPrinterName = "" Then
        MsgBox "プリンタ名が設定されていません。", vbOKOnly
        End
    End If
    
    ' 印刷
    GetRange("賞状氏名").Parent.PrintOut _
        Copies:=1, Collate:=True, IgnorePrintAreas:=False, Preview:=bPreview, _
        ActivePrinter:=sPrinterName

End Sub


'
' 学童大会賞状変数設定
'
Private Sub SetAwardValForGakudo( _
sRaceClass As String, _
sGender As String, _
sDistance As String, _
sStyle As String)
    GetRange("賞状種目区分").Value = sRaceClass & sGender
    GetRange("賞状距離").Value = sDistance
    GetRange("賞状種目").Value = sStyle
End Sub


'
' 市民大会賞状変数設定
'
Private Sub SetAwardValForShimin( _
sRaceClass As String, _
sGender As String, _
sDistance As String, _
sStyle As String, _
sClass As String)
    If sRaceClass = "年齢区分" Then
        GetRange("賞状種目区分").Value = sGender
        GetRange("賞状種目距離区分").Value = sDistance & "Ｍ" & sStyle & "　" & sClass
    Else
        GetRange("賞状種目区分").Value = sRaceClass & sGender
        GetRange("賞状種目距離区分").Value = sDistance & "Ｍ" & sStyle
    End If
    GetRange("賞状大会回数１").Value = GetRange("大会回数").Value
    GetRange("賞状大会回数２").Value = GetRange("大会回数").Value
    GetRange("賞状年").Value = GetRange("大会元号年").Value
    GetRange("賞状月").Value = GetRange("大会月").Value
    GetRange("賞状日").Value = GetRange("大会日").Value

    ' カラム幅の変更
    If sStyle Like "*リレー" Then
        GetRange("賞状氏名").ColumnWidth = 1.13
        GetRange("賞状所属").ColumnWidth = 2.5
    Else
        GetRange("賞状氏名").ColumnWidth = 2.5
        GetRange("賞状所属").ColumnWidth = 1.13
    End If

End Sub

'
' 選手権賞状変数設定
'
Private Sub SetAwardValForSenshuken( _
sGender As String, _
sDistance As String, _
sStyle As String)
    GetRange("賞状性別").Value = sGender
    GetRange("賞状距離").Value = sDistance
    GetRange("賞状種目").Value = sStyle
    
    GetRange("賞状大会回数").Value = GetRange("大会回数").Value
    GetRange("賞状年").Value = GetRange("大会元号年").Value
    GetRange("賞状月").Value = GetRange("大会月").Value
    GetRange("賞状日").Value = GetRange("大会日").Value
End Sub

