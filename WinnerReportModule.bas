Attribute VB_Name = "WinnerReportModule"
Sub 優勝者設定()

    Dim sGameName As String
    sGameName = GetRange("大会名").Value

    Dim sSheetName As String
    sSheetName = GetWinnerSheetName(sGameName)
    Sheets(sSheetName).Select

    Dim sAreaName As String
    sAreaName = GetWinnerAreaName(sGameName)
    
    Dim sMasterName As String
    sMasterName = GetMaster(sGameName)

    Dim sName As String
    Dim sTeam As String
    Dim sType As String
    Dim sTime As String
    
    ' プログラム番号毎
    For Each nProNo In GetAreaKeyData(sMasterName)
        sType = VLookupArea(nProNo, sMasterName, "区分")
        For Each oCell In GetRange("プログラム番号" & CStr(nProNo))
            ' １位の場合
            If oCell.Offset(0, Range("Header順位").Column - Range("HeaderプロNo").Column).Value = 1 Then
                sName = oCell.Offset(0, Range("Header氏名").Column - Range("HeaderプロNo").Column).Value
                sTeam = oCell.Offset(0, Range("Header所属").Column - Range("HeaderプロNo").Column).Value
                If sType = "" Then
                    sType = oCell.Offset(0, Range("Header区分").Column - Range("HeaderプロNo").Column).Value
                End If
                sTime = oCell.Offset(0, Range("Header時間").Column - Range("HeaderプロNo").Column).Value
                
                Call SetWinner(sAreaName, CInt(nProNo), sType, sName, sTeam, sTime)
                Debug.Print "名前：" & sName & "：区分：" & sType & "：時間：" & sTime
            End If
        Next oCell
    Next nProNo

End Sub

'
' 優勝者設定
Sub SetWinner(sAreaName As String, nProNo As Integer, sType As String, sName As String, sTeam As String, sTime As String)
    Dim nTypeCol As Integer
    nTypeCol = GetAreaColumnIndex(sAreaName, "区分")
    Dim nNameCol As Integer
    nNameCol = GetAreaColumnIndex(sAreaName, "氏名")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sAreaName, "所属")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sAreaName, "記録")
    
    ' プログラム番号毎
    For Each oCell In GetAreaKeyData(sAreaName)
        ' 種目番号と区分が一致した場合
        If oCell.Value = nProNo Then
            If oCell.Offset(0, nTypeCol - 1).Value = sType Then
                oCell.Offset(0, nNameCol - 1).Value = sName
                oCell.Offset(0, nTeamCol - 1).Value = sTeam
                oCell.Offset(0, nTimeCol - 1).Value = sTime
            Else
                oCell.Offset(0, nNameCol - 1).Value = ""
                oCell.Offset(0, nTeamCol - 1).Value = ""
                oCell.Offset(0, nTimeCol - 1).Value = ""
            End If
        ElseIf oCell.Value > nProNo Then
            Exit Sub
        End If
    Next oCell

End Sub

'
' 優勝者シート名
'
' sGameName  IN      大会名
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
' sGameName  IN      大会名
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

