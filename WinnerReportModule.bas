Attribute VB_Name = "WinnerReportModule"
Sub �D���Ґݒ�()

    Dim sGameName As String
    sGameName = GetRange("��").Value

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
    
    ' �v���O�����ԍ���
    For Each nProNo In GetAreaKeyData(sMasterName)
        sType = VLookupArea(nProNo, sMasterName, "�敪")
        For Each oCell In GetRange("�v���O�����ԍ�" & CStr(nProNo))
            ' �P�ʂ̏ꍇ
            If oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value = 1 Then
                sName = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                sTeam = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                If sType = "" Then
                    sType = oCell.Offset(0, Range("Header�敪").Column - Range("Header�v��No").Column).Value
                End If
                sTime = oCell.Offset(0, Range("Header����").Column - Range("Header�v��No").Column).Value
                
                Call SetWinner(sAreaName, CInt(nProNo), sType, sName, sTeam, sTime)
                Debug.Print "���O�F" & sName & "�F�敪�F" & sType & "�F���ԁF" & sTime
            End If
        Next oCell
    Next nProNo

End Sub

'
' �D���Ґݒ�
Sub SetWinner(sAreaName As String, nProNo As Integer, sType As String, sName As String, sTeam As String, sTime As String)
    Dim nTypeCol As Integer
    nTypeCol = GetAreaColumnIndex(sAreaName, "�敪")
    Dim nNameCol As Integer
    nNameCol = GetAreaColumnIndex(sAreaName, "����")
    Dim nTeamCol As Integer
    nTeamCol = GetAreaColumnIndex(sAreaName, "����")
    Dim nTimeCol As Integer
    nTimeCol = GetAreaColumnIndex(sAreaName, "�L�^")
    
    ' �v���O�����ԍ���
    For Each oCell In GetAreaKeyData(sAreaName)
        ' ��ڔԍ��Ƌ敪����v�����ꍇ
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
' �D���҃V�[�g��
'
' sGameName  IN      ��
'
Function GetWinnerSheetName(sGameName As String)
    If sGameName = "���{��I�茠���j���" Then
        GetWinnerSheetName = "�I�茠���D����"
    ElseIf sGameName = "���{��s���̈���" Then
        GetWinnerSheetName = "�s�����D����"
    Else
        GetWinnerSheetName = "�w���}�X�^�[�Y�D����"
    End If
End Function

'
' �D���Ҕ͈͖�
'
' sGameName  IN      ��
'
Function GetWinnerAreaName(sGameName As String)
    If sGameName = "���{��I�茠���j���" Then
        GetWinnerAreaName = "�I�茠���D����"
    ElseIf sGameName = "���{��s���̈���" Then
        GetWinnerAreaName = "�s�����D����"
    Else
        GetWinnerAreaName = "�w�}���D����"
    End If
End Function

