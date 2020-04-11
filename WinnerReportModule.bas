Attribute VB_Name = "WinnerReportModule"
Sub DÒÝè()

    Dim sGameName As String
    sGameName = GetRange("åï¼").Value

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
    
    ' vOÔ
    For Each nProNo In GetAreaKeyData(sMasterName)
        sType = VLookupArea(nProNo, sMasterName, "æª")
        For Each oCell In GetRange("vOÔ" & CStr(nProNo))
            ' PÊÌê
            If oCell.Offset(0, Range("HeaderÊ").Column - Range("HeadervNo").Column).Value = 1 Then
                sName = oCell.Offset(0, Range("Header¼").Column - Range("HeadervNo").Column).Value
                sTeam = oCell.Offset(0, Range("Header®").Column - Range("HeadervNo").Column).Value
                If sType = "" Then
                    sType = oCell.Offset(0, Range("Headeræª").Column - Range("HeadervNo").Column).Value
                End If
                sTime = oCell.Offset(0, Range("HeaderÔ").Column - Range("HeadervNo").Column).Value
                
                Call SetWinner(sAreaName, CInt(nProNo), sType, sName, sTeam, sTime)
                Debug.Print "¼OF" & sName & "FæªF" & sType & "FÔF" & sTime
            End If
        Next oCell
    Next nProNo

End Sub

'
' DÒÝè
Sub SetWinner(sAreaName As String, nProNo As Integer, sType As String, sName As String, sTeam As String, sTime As String)
    Dim nTypeCol As Integer
    nTypeCol = GetAreaColumnIndex(sAreaName, "æª")
    Dim nNameCol As Integer
    nNameCol = GetAreaColumnIndex(sAreaName, "¼")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sAreaName, "®")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sAreaName, "L^")
    
    ' vOÔ
    For Each oCell In GetAreaKeyData(sAreaName)
        ' íÚÔÆæªªêvµ½ê
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
' DÒV[g¼
'
' sGameName  IN      åï¼
'
Function GetWinnerSheetName(sGameName As String)
    If sGameName = "¡{êIè jåï" Then
        GetWinnerSheetName = "Iè åïDÒ"
    ElseIf sGameName = "¡{ês¯Ìçåï" Then
        GetWinnerSheetName = "s¯åïDÒ"
    Else
        GetWinnerSheetName = "w¶}X^[YDÒ"
    End If
End Function

'
' DÒÍÍ¼
'
' sGameName  IN      åï¼
'
Function GetWinnerAreaName(sGameName As String)
    If sGameName = "¡{êIè jåï" Then
        GetWinnerAreaName = "Iè åïDÒ"
    ElseIf sGameName = "¡{ês¯Ìçåï" Then
        GetWinnerAreaName = "s¯åïDÒ"
    Else
        GetWinnerAreaName = "w}åïDÒ"
    End If
End Function

