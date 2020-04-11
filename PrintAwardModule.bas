Attribute VB_Name = "PrintAwardModule"
'
' 賞状印刷
'
' 指定したレースNoに存在するProNoの賞状を印刷する
'
Sub 賞状印刷()

    ' イベント発生を抑制
    Call EventChange(False)

    Dim oWorkSheet As Worksheet
    Set oWorkSheet = ActiveSheet

    ' レース番号
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
                ' 賞状印刷対象か確認
                If CheckTarget(nProNo) Then
                    ' 賞状を印刷
                    Call PrintAward(nProNo)
                End If
                oProNo.Add nProNo, 1
            End If
        Next vRaceNo
    End If

    oWorkSheet.Activate

    ' イベント発生を再開
    Call EventChange(True)

End Sub

'
' 印刷対象かを確認する
'
' nProNo            IN      プロNo
'
Function CheckTarget(nProNo As Integer) As Boolean

    ' 種目の区分を取得
    Dim sGameType
    sGameType = Application.WorksheetFunction.VLookup(nProNo, Range("学マ種目区分"), 6, False)
    
    If sGameType = "学童" Or sGameType = "学童リレー" Then
        CheckTarget = True
    Else
        CheckTarget = False
    End If

End Function

'
' 賞状を印刷する
'
' 指定したプロNoの中で1位～3位の賞状を印刷する
'
' nProNo            IN      プロNo
'
Sub PrintAward(nProNo As Integer)

    Dim sName As String
    sName = "プログラム番号" & Trim(Str(nProNo))
    
    Dim oWorkSheet As Worksheet
    Set oWorkSheet = GetRange(sName).Parent
    
    Dim sRaceType As String ' 区分
    sRaceType = Application.WorksheetFunction.VLookup(nProNo, Range("学マ種目区分"), 2, False)
    Dim sGender As String ' 性別
    sGender = Application.WorksheetFunction.VLookup(nProNo, Range("学マ種目区分"), 3, False)
    Dim sDistance As String ' 距離
    sDistance = Application.WorksheetFunction.VLookup(nProNo, Range("学マ種目区分"), 4, False)
    Dim sStyle As String ' 種目名
    sStyle = Application.WorksheetFunction.VLookup(nProNo, Range("学マ種目区分"), 5, False)
    
    Dim nOrder As Integer
    For Each vProNo In GetRange(sName)
        nOrder = Val(vProNo.Offset(0, GetRange("Prog順位").Column - vProNo.Column).Value)
        If nOrder >= 1 And nOrder <= 3 Then
            Call PrintAwardByLine(oWorkSheet, vProNo.Row, sRaceType, sGender, sDistance, sStyle)
        End If
    Next vProNo
    
End Sub

'
' 行指定で賞状を印刷する
'
' 指定した行のレコードを印刷する
'
' oWorkSheet        IN      ワークシート
' nRow              IN      行番号
' sRaceType         IN      種目区分
' sGender           IN      性別
' sDistance         IN      距離
' sStyle            IN      種目名
'
Sub PrintAwardByLine(oWorkSheet As Worksheet, _
nRow As Integer, _
sRaceType As String, _
sGender As String, _
sDistance As String, _
sStyle As String)
   
    GetRange("賞状種目区分").Value = sRaceType & sGender
    GetRange("賞状距離").Value = sDistance
    GetRange("賞状種目").Value = sStyle
    GetRange("賞状順位").Value = oWorkSheet.Cells(nRow, Range("Prog順位").Column).Value
    GetRange("賞状タイム").Value = oWorkSheet.Cells(nRow, Range("Prog時間").Column).Value
    GetRange("賞状大会新").Value = oWorkSheet.Cells(nRow, Range("Prog備考").Column).Value
    GetRange("賞状氏名").Value = oWorkSheet.Cells(nRow, Range("Prog氏名").Column).Value
    GetRange("賞状所属").Value = oWorkSheet.Cells(nRow, Range("Prog所属").Column).Value

    If GetRange("賞状タイム").Value >= 10000 Then
        GetRange("賞状タイム").NumberFormatLocal = "#""分""##""秒""##"
    Else
        GetRange("賞状タイム").NumberFormatLocal = "##""秒""##"
    End If

    ' 印刷
    GetRange("賞状種目区分").Parent.Activate
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, Preview:=True

End Sub

Sub 賞状名前定義(Optional sValue As String = "")
    Call 学マ賞状名前定義
End Sub

Sub 学マ賞状名前定義(Optional sValue As String = "")
    Sheets("学マ賞状").Select
    ActiveSheet.Unprotect

    ' 名前をすべて削除
    Call DeleteName("賞状*")

    Call SetName("賞状種目区分", "$C$9")
    Call SetName("賞状距離", "$G$9")
    Call SetName("賞状種目", "$L$9")
    Call SetName("賞状順位", "$A$13")
    Call SetName("賞状タイム", "$L$14")
    Call SetName("賞状大会新", "$S$14")
    Call SetName("賞状氏名", "$C$20")
    Call SetName("賞状所属", "$C$24")
 
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
 
End Sub

